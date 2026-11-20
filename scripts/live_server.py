"""Local ONLYOFFICE integration. Run inside the isolated editor Docker network."""
from __future__ import annotations

import base64
import hashlib
import hmac
import io
import json
import os
from pathlib import Path
import re
import secrets
import signal
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.error import HTTPError
from urllib.parse import parse_qs, urlsplit
from urllib.request import Request, urlopen, HTTPRedirectHandler, build_opener
import zipfile

DATA = Path(os.environ.get('OFFICE_DATA', '/data'))
ASSETS = Path(__file__).resolve().parent.parent / 'assets'
API_TOKEN = os.environ.get('OFFICE_API_TOKEN', '')
JWT_SECRET = os.environ.get('JWT_SECRET', '')
DOCS_INTERNAL = os.environ.get('DOCS_INTERNAL', 'http://office-docs')
DOCS_PUBLIC = os.environ.get('DOCS_PUBLIC', 'http://127.0.0.1:18821')
PUBLIC = os.environ.get('OFFICE_PUBLIC', 'http://127.0.0.1:18820')
GUID = 'asc.{0F5AB88D-13DB-4D61-AC54-33A9005228EB}'
MAX_BYTES = 100 * 1024 * 1024
LOCK = threading.RLock()
COMMANDS: dict[str, list[dict]] = {}
SAVE_LOCKS: dict[str, threading.RLock] = {}
KINDS = {'.docx': 'word', '.xlsx': 'cell', '.pptx': 'slide'}


def read_json(path, default=None):
    try:
        return json.loads(path.read_text(encoding='utf-8'))
    except FileNotFoundError:
        return default


def write_json(path, value):
    temporary = path.with_name(path.name + '.' + secrets.token_hex(5) + '.tmp')
    temporary.write_text(json.dumps(value, ensure_ascii=False), encoding='utf-8')
    os.replace(temporary, path)


def b64(data):
    return base64.urlsafe_b64encode(data).decode().rstrip('=')


def sign(value):
    header = b64(b'{"alg":"HS256","typ":"JWT"}')
    payload = b64(json.dumps(value, separators=(',', ':')).encode())
    raw = header + '.' + payload
    return raw + '.' + b64(hmac.digest(JWT_SECRET.encode(), raw.encode(), 'sha256'))


def verify(token):
    parts = token.split('.')
    if len(parts) != 3:
        raise ValueError('Invalid signed callback')
    expected = b64(hmac.digest(JWT_SECRET.encode(), (parts[0] + '.' + parts[1]).encode(), 'sha256'))
    if not hmac.compare_digest(expected, parts[2]):
        raise ValueError('Invalid callback signature')
    header = json.loads(base64.urlsafe_b64decode(parts[0] + '=' * (-len(parts[0]) % 4)))
    if header.get('alg') != 'HS256':
        raise ValueError('Unsupported signature algorithm')
    return json.loads(base64.urlsafe_b64decode(parts[1] + '=' * (-len(parts[1]) % 4)))


def session_path(identifier):
    if not re.fullmatch(r'[a-f0-9]{32}', identifier):
        raise ValueError('Invalid session identifier')
    path = DATA / 'sessions' / identifier
    if not (path / 'session.json').is_file():
        raise FileNotFoundError('Unknown editing session')
    return path


def load_session(identifier):
    return read_json(session_path(identifier) / 'session.json')


def status(identifier):
    root = session_path(identifier)
    info = read_json(root / 'session.json')
    return {'id': identifier, 'title': info['title'], 'kind': info['kind'],
            'bridge': read_json(root / 'bridge.json', {}),
            'revision': read_json(root / 'revision.json', {'number': 0}),
            'save': read_json(root / 'save-status.json', {'state': 'waiting'}),
            'editor': read_json(root / 'editor-status.json', {})}


def editor_config(identifier):
    info = load_session(identifier)
    cfg = {
        'document': {'fileType': info['extension'].lstrip('.'), 'key': info['key'],
                     'title': info['title'],
                     'url': f'http://office-api:8000/file/{identifier}?token={info["fetch_token"]}',
                     'permissions': {'edit': True, 'download': True, 'print': True}},
        'documentType': {'word': 'word', 'cell': 'cell', 'slide': 'slide'}[info['kind']],
        'editorConfig': {
            'mode': 'edit', 'lang': 'zh-CN',
            'callbackUrl': f'http://office-api:8000/callback/{identifier}',
            'user': {'id': 'local-owner', 'name': '本机用户'},
            'customization': {'forcesave': True, 'autosave': True,
                              'goback': {'url': PUBLIC, 'text': '返回会话'}},
            'plugins': {'autostart': [GUID],
                        'pluginsData': [f'{PUBLIC}/plugin/{identifier}/{info["ui_token"]}/config.json'],
                        'options': {GUID: {'session': identifier, 'token': info['ui_token'],
                                            'base': PUBLIC}}}},
        'height': '100%', 'width': '100%', 'type': 'desktop'}
    cfg['token'] = sign(cfg)
    return cfg


def force_save(identifier):
    info = load_session(identifier)
    with LOCK:
        root = session_path(identifier)
        generation = read_json(root / 'save-generation.json', {'number': 0})['number'] + 1
        write_json(root / 'save-generation.json', {'number': generation})
    command = {'c': 'forcesave', 'key': info['key'], 'userdata': str(generation)}
    command['token'] = sign(command)
    request = Request(DOCS_INTERNAL + '/coauthoring/CommandService.ashx',
                      data=json.dumps(command).encode(), headers={'Content-Type': 'application/json'})
    with urlopen(request, timeout=15) as response:
        result = json.load(response)
        result['generation'] = generation
        return result


def store_callback(identifier, body, bearer):
    with LOCK:
        save_lock = SAVE_LOCKS.setdefault(identifier, threading.RLock())
    with save_lock:
        return store_callback_serial(identifier, body, bearer)


class NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, request, fp, code, msg, headers, newurl):
        return None


def store_callback_serial(identifier, body, bearer):
    signed = verify(bearer or body.get('token', ''))
    payload = signed.get('payload', signed)
    # Use signed fields, never unsigned body fields when JWT is present.
    if 'status' not in payload:
        raise ValueError('Callback signature did not cover its status')
    info = load_session(identifier)
    if payload.get('key') != info['key']:
        raise ValueError('Callback document key mismatch')
    state = int(payload['status'])
    if state not in (2, 6):
        if state in (3, 7):
            write_json(session_path(identifier) / 'editor-status.json',
                       {'state': 'save_error', 'status': state, 'time': time.time()})
        return {'error': 0}
    root = session_path(identifier)
    previous = read_json(root / 'revision.json', {'number': 0, 'generation': 0})
    generation = int(payload.get('userdata', 0))
    if generation and generation <= previous.get('generation', 0):
        return {'error': 0}
    if not generation and previous.get('generation', 0):
        # Unsequenced autosave must not roll back an acknowledged explicit save.
        # The normal user workflow uses the integrated explicit Save button.
        write_json(root / 'editor-status.json', {'state': 'explicit_save_required', 'time': time.time()})
        return {'error': 1}
    remote = urlsplit(payload['url'])
    allowed = {urlsplit(DOCS_INTERNAL).netloc, urlsplit(DOCS_PUBLIC).netloc}
    if remote.scheme != 'http' or remote.netloc not in allowed or not remote.path.startswith('/cache/'):
        raise ValueError('Unexpected saved-document location')
    location = DOCS_INTERNAL + remote.path + ('?' + remote.query if remote.query else '')
    # Saved files are always fetched from the fixed local document service.
    with build_opener(NoRedirect).open(location, timeout=30) as response:
        if response.geturl() != location:
            raise ValueError('Saved-document redirect rejected')
        data = response.read(MAX_BYTES + 1)
    if len(data) > MAX_BYTES:
        raise ValueError('Saved document exceeds 100MB')
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        required = {'word': 'word/document.xml', 'cell': 'xl/workbook.xml',
                    'slide': 'ppt/presentation.xml'}[info['kind']]
        if required not in archive.namelist():
            raise ValueError('Saved output format mismatch')
    with LOCK:
        content_hash = hashlib.sha256(data).hexdigest()
        if content_hash == previous.get('sha256'):
            previous['generation'] = max(previous.get('generation', 0), generation)
            previous['time'] = time.time()
            write_json(root / 'revision.json', previous)
            return {'error': 0}
        number = previous['number'] + 1
        filename = f'revision-{number:06d}{info["extension"]}'
        temporary = root / (filename + '.tmp')
        temporary.write_bytes(data)
        os.replace(temporary, root / filename)
        write_json(root / 'revision.json', {'number': number, 'file': filename,
                   'sha256': content_hash, 'generation': generation, 'time': time.time()})
    return {'error': 0}


def validate_command(kind, body):
    op = body.get('op')
    allowed = {'inspect', 'word.append', 'word.replace', 'sheet.set', 'sheet.read', 'slide.set_text'}
    if op not in allowed:
        raise ValueError('Unsupported operation')
    if op != 'inspect' and not op.startswith({'word': 'word.', 'cell': 'sheet.', 'slide': 'slide.'}[kind]):
        raise ValueError('Operation does not match document type')
    if len(json.dumps(body)) > 1_000_000:
        raise ValueError('Command is too large')
    if op in ('word.append', 'slide.set_text') and not isinstance(body.get('text'), str):
        raise ValueError('text must be a string')
    if op == 'word.replace' and (not isinstance(body.get('search'), str) or not body['search'] or
                                  not isinstance(body.get('replace'), str)):
        raise ValueError('search and replace strings required')
    if op.startswith('sheet.') and not re.fullmatch(r'[A-Za-z]{1,3}[1-9][0-9]{0,6}(:[A-Za-z]{1,3}[1-9][0-9]{0,6})?', body.get('range', '')):
        raise ValueError('A1-style range is required')
    if op == 'sheet.set' and not isinstance(body.get('values'), (list, str, int, float, bool)):
        raise ValueError('values must be a scalar or array')
    if op == 'slide.set_text' and any(not isinstance(body.get(k), int) or body[k] < 0 for k in ('slide', 'shape')):
        raise ValueError('slide and shape must be zero-based nonnegative indexes')


class Handler(BaseHTTPRequestHandler):
    server_version = 'MicrosoftWpsLive/1'

    def log_message(self, format, *args):
        # Do not log capabilities, signed URLs, or document content.
        pass

    def send(self, value, content_type='application/json', code=200):
        data = value if isinstance(value, bytes) else (value.encode() if isinstance(value, str)
                                                     else json.dumps(value, ensure_ascii=False).encode())
        self.send_response(code)
        self.send_header('Content-Type', content_type + ('; charset=utf-8' if 'text' in content_type or 'json' in content_type else ''))
        self.send_header('Content-Length', str(len(data)))
        self.send_header('Cache-Control', 'no-store')
        self.send_header('Referrer-Policy', 'no-referrer')
        self.send_header('X-Content-Type-Options', 'nosniff')
        if urlsplit(self.path).path.startswith('/plugin/'):
            self.send_header('Access-Control-Allow-Origin', DOCS_PUBLIC)
            self.send_header('Vary', 'Origin')
        self.end_headers()
        self.wfile.write(data)

    def body(self):
        size = int(self.headers.get('Content-Length', '0'))
        if size > 1_000_000:
            raise ValueError('Request too large')
        return json.loads(self.rfile.read(size) or b'{}')

    def authorized(self, identifier=None, fetch=False):
        supplied = self.headers.get('Authorization', '').removeprefix('Bearer ')
        if supplied and hmac.compare_digest(supplied, API_TOKEN):
            return True
        if identifier:
            info = load_session(identifier)
            if fetch:
                supplied = parse_qs(urlsplit(self.path).query).get('token', [''])[0]
                return bool(supplied) and hmac.compare_digest(supplied, info['fetch_token'])
            return bool(supplied) and hmac.compare_digest(supplied, info['ui_token'])
        return False

    def do_GET(self):
        self.dispatch('GET')

    def do_POST(self):
        self.dispatch('POST')

    def dispatch(self, method):
        try:
            self.route(method)
        except (ValueError, KeyError, json.JSONDecodeError) as error:
            self.send({'error': str(error)}, code=400)
        except FileNotFoundError:
            self.send({'error': 'Not found'}, code=404)
        except Exception as error:
            self.send({'error': type(error).__name__}, code=500)

    def route(self, method):
        path = urlsplit(self.path).path
        parts = path.strip('/').split('/')
        if path == '/health':
            return self.send({'ok': True})
        if path == '/':
            return self.send((ASSETS / 'home.html').read_bytes(), 'text/html')
        if len(parts) == 2 and parts[0] == 'edit':
            session_path(parts[1])
            content = (ASSETS / 'editor.html').read_text(encoding='utf-8')
            return self.send(content.replace('__SESSION__', parts[1]).replace('__DOCS__', DOCS_PUBLIC), 'text/html')
        if len(parts) == 4 and parts[0] == 'plugin':
            identifier, capability, item = parts[1:]
            info = load_session(identifier)
            if not hmac.compare_digest(capability, info['ui_token']):
                return self.send({'error': 'Unauthorized'}, code=401)
            if item == 'config.json':
                return self.send({'name': 'microsoft-wps-live Agent Bridge', 'guid': GUID, 'version': '1.0.0',
                                  'variations': [{'url': 'index.html', 'description': 'Local live editing bridge',
                                   'isVisual': False, 'isViewer': False,
                                   'EditorsSupport': ['word', 'cell', 'slide'], 'initDataType': 'none',
                                   'initData': '', 'buttons': []}]})
            if item == 'index.html':
                content = (ASSETS / 'plugin.html').read_text(encoding='utf-8')
                return self.send(content.replace('__SESSION__', identifier)
                                 .replace('__CAPABILITY__', info['ui_token']), 'text/html')
            if item == 'bridge.js':
                return self.send((ASSETS / 'bridge.js').read_bytes(), 'application/javascript')
            if item == 'sdk.js':
                with urlopen(DOCS_INTERNAL + '/sdkjs-plugins/v1/plugins.js', timeout=15) as response:
                    return self.send(response.read(2_000_000), 'application/javascript')
        if len(parts) == 2 and parts[0] == 'file':
            if not self.authorized(parts[1], fetch=True):
                return self.send({'error': 'Unauthorized'}, code=401)
            info = load_session(parts[1])
            root = session_path(parts[1])
            revision = read_json(root / 'revision.json', {})
            filename = revision.get('file', 'input' + info['extension'])
            if filename != 'input' + info['extension'] and not re.fullmatch(r'revision-[0-9]{6}\.(docx|xlsx|pptx)', filename):
                raise ValueError('Invalid session file')
            return self.send((root / filename).read_bytes(), 'application/octet-stream')
        if len(parts) == 2 and parts[0] == 'callback' and method == 'POST':
            bearer = self.headers.get('Authorization', '').removeprefix('Bearer ')
            return self.send(store_callback(parts[1], self.body(), bearer))
        if path == '/api/sessions' and method == 'GET':
            if not self.authorized():
                return self.send({'error': 'Unauthorized'}, code=401)
            return self.send([status(p.name) for p in (DATA / 'sessions').iterdir()
                              if (p / 'session.json').is_file()])
        if len(parts) >= 4 and parts[:2] == ['api', 'session']:
            identifier, action = parts[2:4]
            if not self.authorized(identifier):
                return self.send({'error': 'Unauthorized'}, code=401)
            root = session_path(identifier)
            if action == 'config' and method == 'GET':
                return self.send(editor_config(identifier))
            if action == 'status' and method == 'GET':
                return self.send(status(identifier))
            if action == 'save' and method == 'POST':
                self.body()
                return self.send(force_save(identifier))
            if action == 'ready' and method == 'POST':
                self.body()
                write_json(root / 'editor-status.json', {'state': 'ready', 'time': time.time()})
                return self.send({'ok': True})
            if action == 'poll' and method == 'GET':
                with LOCK:
                    owner = self.headers.get('X-Office-Bridge', '')
                    if not re.fullmatch(r'[a-zA-Z0-9-]{16,64}', owner):
                        raise ValueError('Bridge instance identifier required')
                    bridge = read_json(root / 'bridge.json', {})
                    if bridge.get('owner') != owner and time.time() - bridge.get('time', 0) < 10:
                        return self.send({'command': None, 'state': 'another_bridge_active'})
                    write_json(root / 'bridge.json', {'state': 'online', 'owner': owner, 'time': time.time()})
                    active = next((c for c in COMMANDS.get(identifier, [])
                                   if c['state'] in ('running', 'unknown')), None)
                    if active:
                        if time.time() - active.get('started', time.time()) > 20:
                            active['state'] = 'unknown'
                        return self.send({'command': None, 'state': active['state']})
                    pending = next((c for c in COMMANDS.get(identifier, []) if c['state'] == 'queued'), None)
                    if pending:
                        pending['state'] = 'running'
                        pending['started'] = time.time()
                        pending['owner'] = owner
                        return self.send({'command': pending})
                    return self.send({'command': None})
            if action == 'complete' and method == 'POST':
                body = self.body()
                with LOCK:
                    command = next(c for c in COMMANDS.get(identifier, []) if c['id'] == body['id'])
                    if command.get('owner') != self.headers.get('X-Office-Bridge'):
                        raise ValueError('Command belongs to another bridge')
                    if command['state'] in ('applied', 'failed'):
                        return self.send({'ok': True, 'already_completed': True})
                    if command['state'] not in ('running', 'unknown'):
                        raise ValueError('Command is not running')
                    command['state'] = 'applied' if body.get('ok') else 'failed'
                    command['result'] = body.get('result')
                    command['finished'] = time.time()
                return self.send({'ok': True})
            if action == 'commands' and method == 'POST':
                body = self.body()
                info = load_session(identifier)
                validate_command(info['kind'], body)
                bridge = read_json(root / 'bridge.json', {})
                if time.time() - bridge.get('time', 0) > 10:
                    return self.send({'error': 'Live editor bridge is offline. Open the editor first.'}, code=409)
                command = {'id': secrets.token_hex(16), 'state': 'queued', 'body': body, 'created': time.time()}
                with LOCK:
                    COMMANDS.setdefault(identifier, []).append(command)
                return self.send(command)
            if action == 'commands' and len(parts) == 5 and method == 'GET':
                with LOCK:
                    command = next(c for c in COMMANDS.get(identifier, []) if c['id'] == parts[4])
                    return self.send(command)
        self.send({'error': 'Not found'}, code=404)


if __name__ == '__main__':
    if not API_TOKEN or not JWT_SECRET:
        raise SystemExit('Private API token and JWT secret are required')
    (DATA / 'sessions').mkdir(parents=True, exist_ok=True)
    server = ThreadingHTTPServer(('0.0.0.0', 8000), Handler)
    signal.signal(signal.SIGTERM, lambda *_: threading.Thread(target=server.shutdown, daemon=True).start())
    try:
        server.serve_forever()
    finally:
        server.server_close()
