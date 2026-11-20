"""Shared local CLI for opening, editing and saving Office live sessions."""
from __future__ import annotations
import argparse
import atexit
import hashlib
import json
import os
from pathlib import Path
import re
import secrets
import shutil
import socket
import subprocess
import sys
import time
from urllib.request import Request, urlopen
import zipfile

ROOT = Path(__file__).resolve().parent.parent
CONFIG = ROOT / 'config'
DATA = ROOT / 'data'
BINDINGS = CONFIG / 'bindings'
KINDS = {'.docx': 'word', '.xlsx': 'cell', '.pptx': 'slide'}
BASE = 'http://127.0.0.1:18820'


def read(path, default=None):
    try:
        return json.loads(path.read_text(encoding='utf-8'))
    except FileNotFoundError:
        return default


def write(path, value):
    temporary = path.with_name(path.name + '.' + secrets.token_hex(5) + '.tmp')
    temporary.write_text(json.dumps(value, ensure_ascii=False), encoding='utf-8')
    os.replace(temporary, path)


def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def linux_path(path):
    text = str(path.resolve())
    if re.match(r'^[A-Za-z]:\\', text):
        return '/mnt/' + text[0].lower() + '/' + text[3:].replace('\\', '/')
    return text


def settings():
    result = read(CONFIG / 'settings.json')
    if not result:
        raise RuntimeError('Run init and configure the local runtime first.')
    return result


def api(action, body=None):
    cfg = settings()
    request = Request(BASE + action, data=None if body is None else json.dumps(body).encode(),
                      headers={'Authorization': 'Bearer ' + cfg['api_token'],
                               'Content-Type': 'application/json'})
    with urlopen(request, timeout=20) as response:
        return json.load(response)


def initialize():
    for name in ('config/bindings', 'data/sessions', 'data/backups', 'logs', 'downloads'):
        (ROOT / name).mkdir(parents=True, exist_ok=True)
    cfg = read(CONFIG / 'settings.json')
    if not cfg:
        cfg = {'api_token': secrets.token_urlsafe(48), 'jwt_secret': secrets.token_urlsafe(48),
               'distro': 'OfficeLiveEditor'}
        write(CONFIG / 'settings.json', cfg)
    private = CONFIG / 'private.env'
    private.write_text('OFFICE_API_TOKEN=' + cfg['api_token'] + '\nJWT_SECRET=' + cfg['jwt_secret'] +
                       '\nJWT_ENABLED=true\nJWT_HEADER=Authorization\nALLOW_PRIVATE_IP_ADDRESS=true\n', encoding='utf-8')
    project = linux_path(ROOT)
    compose = {'name': 'office-live-editor', 'services': {
        'office-api': {'image': 'python:3.12-alpine', 'command': ['python', '/app/scripts/live_server.py'],
                       'env_file': [project + '/config/private.env'], 'restart': 'unless-stopped',
                       'ports': ['127.0.0.1:18820:8000'],
                       'volumes': [project + '/scripts:/app/scripts:ro', project + '/assets:/app/assets:ro',
                                   project + '/data:/data'], 'networks': ['office-local']},
        'office-docs': {'image': 'onlyoffice/documentserver:10.0.0', 'restart': 'unless-stopped',
                        'env_file': [project + '/config/private.env'], 'shm_size': '1gb',
                        'ports': ['127.0.0.1:18821:80'], 'networks': ['office-local'],
                        'volumes': ['docs-data:/var/www/onlyoffice/Data',
                                    'docs-db:/var/lib/postgresql', 'docs-cache:/var/lib/onlyoffice']}},
        'networks': {'office-local': {'driver': 'bridge',
                     'driver_opts': {'com.docker.network.bridge.host_binding_ipv4': '127.0.0.1'}}},
        'volumes': {'docs-data': {}, 'docs-db': {}, 'docs-cache': {}}}
    write(CONFIG / 'compose.json', compose)
    return {'configured': True, 'root': str(ROOT)}


def compose(*args):
    cfg = settings()
    command = ['wsl', '-d', cfg['distro'], '-u', 'root', '--', 'docker', 'compose',
               '-f', linux_path(CONFIG / 'compose.json'), *args]
    subprocess.run(command, check=True)


def open_session(source, new=False):
    source = source.resolve(strict=True)
    extension = source.suffix.lower()
    if extension not in KINDS:
        raise ValueError('Live editing supports DOCX, XLSX and PPTX.')
    if source.stat().st_size > 100 * 1024 * 1024:
        raise ValueError('Maximum document size is 100MB.')
    with zipfile.ZipFile(source) as archive:
        if any('vbaproject' in s.lower() or 'macrosheet' in s.lower() for s in archive.namelist()):
            raise ValueError('Macro-bearing documents are not supported.')
    api('/health')
    bindings = {p.stem: read(p) for p in BINDINGS.glob('*.json')}
    if not new:
        for identifier, binding in bindings.items():
            if binding['source'] == str(source):
                if sha(source) != binding['sha256']:
                    raise RuntimeError('The file changed outside this editing session. Resolve the conflict before reopening.')
                info = read(DATA / 'sessions' / identifier / 'session.json')
                return {'session': identifier, 'url': BASE + '/edit/' + identifier + '#token=' + info['ui_token'],
                        'source': str(source), 'reused': True}
    identifier = secrets.token_hex(16)
    folder = DATA / 'sessions' / identifier
    folder.mkdir()
    original_hash = sha(source)
    shutil.copyfile(source, folder / ('input' + extension))
    if sha(source) != original_hash:
        raise RuntimeError('Source changed during session creation; no session was published.')
    info = {'id': identifier, 'title': source.name, 'extension': extension, 'kind': KINDS[extension],
            'key': 'local-' + identifier, 'fetch_token': secrets.token_urlsafe(36),
            'ui_token': secrets.token_urlsafe(36), 'created': time.time()}
    write(BINDINGS / (identifier + '.json'), {'source': str(source), 'sha256': original_hash, 'revision': 0})
    write(folder / 'session.json', info)
    return {'session': identifier, 'url': BASE + '/edit/' + identifier + '#token=' + info['ui_token'],
            'source': str(source), 'reused': False}


def save_revision(identifier, binding):
    folder = DATA / 'sessions' / identifier
    revision = read(folder / 'revision.json', {'number': 0})
    if (revision['number'] <= binding.get('revision', 0) and
            revision.get('generation', 0) <= binding.get('generation', 0)):
        return
    state = folder / 'save-status.json'
    source = Path(binding['source'])
    if not source.is_file() or sha(source) != binding['sha256']:
        write(state, {'state': 'conflict', 'revision': revision['number'], 'time': time.time()})
        return
    filename = revision['file']
    if not re.fullmatch(r'revision-[0-9]{6}\.(docx|xlsx|pptx)', filename):
        raise ValueError('Invalid staged revision filename')
    staged = folder / filename
    if staged.is_symlink() or staged.resolve().parent != folder.resolve():
        raise ValueError('Staged revision escaped its session directory')
    if sha(staged) != revision['sha256']:
        raise ValueError('Staged revision checksum mismatch')
    if revision['number'] <= binding.get('revision', 0) and sha(source) == revision['sha256']:
        binding['generation'] = revision.get('generation', 0)
        write(state, {'state': 'saved', 'revision': revision['number'], 'generation': binding['generation'],
                      'sha256': binding['sha256'], 'time': time.time()})
        return
    backup = DATA / 'backups' / (identifier + '-' + str(revision['number']) + source.suffix)
    if not backup.exists():
        shutil.copyfile(source, backup)
    temporary = source.with_name('.' + source.name + '.office-live-' + secrets.token_hex(5) + '.tmp')
    try:
        shutil.copyfile(staged, temporary)
        # Recheck immediately before replacement to catch external changes.
        if sha(source) != binding['sha256']:
            write(state, {'state': 'conflict', 'revision': revision['number'], 'time': time.time()})
            return
        os.replace(temporary, source)
        binding['sha256'] = sha(source)
        binding['revision'] = revision['number']
        binding['generation'] = revision.get('generation', 0)
        write(state, {'state': 'saved', 'revision': revision['number'], 'sha256': binding['sha256'],
                      'generation': binding['generation'], 'time': time.time()})
    finally:
        if temporary.exists():
            temporary.unlink()


def watch():
    lockfile = (CONFIG / 'watcher.lock').open('a+b')
    if os.name == 'nt':
        import msvcrt
        lockfile.seek(0)
        if lockfile.read(1) == b'':
            lockfile.write(b'0'); lockfile.flush()
        lockfile.seek(0)
        try:
            msvcrt.locking(lockfile.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError:
            return
    else:
        import fcntl
        try:
            fcntl.flock(lockfile, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            return
    keepalive = None
    if os.name == 'nt':
        keepalive = subprocess.Popen(['wsl', '-d', settings()['distro'], '-u', 'root', '--', 'sleep', 'infinity'],
                                     stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                                     stderr=subprocess.DEVNULL, creationflags=subprocess.CREATE_NO_WINDOW)
        atexit.register(lambda: keepalive.terminate() if keepalive.poll() is None else None)
    while not (CONFIG / 'stop-watcher').exists():
        for binding_file in BINDINGS.glob('*.json'):
            identifier, binding = binding_file.stem, read(binding_file)
            try:
                save_revision(identifier, binding)
                write(binding_file, binding)
            except Exception as error:
                write(DATA / 'sessions' / identifier / 'save-status.json',
                      {'state': 'error', 'error': str(error), 'time': time.time()})
        write(CONFIG / 'watcher.json', {'pid': os.getpid(), 'time': time.time()})
        time.sleep(0.7)
    if keepalive is not None and keepalive.poll() is None:
        keepalive.terminate()
        keepalive.wait(timeout=5)


def start_watcher():
    stop = CONFIG / 'stop-watcher'
    if stop.exists():
        stop.unlink()
    heartbeat = read(CONFIG / 'watcher.json', {})
    if time.time() - heartbeat.get('time', 0) < 5:
        return
    with (ROOT / 'logs' / 'save-watcher.log').open('ab') as log:
        subprocess.Popen([sys.executable, str(Path(__file__).resolve()), '_watch'],
                         stdout=log, stderr=log,
                         creationflags=(subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP)
                         if os.name == 'nt' else 0)


def reachable_ports():
    ready = set()
    for port, path in ((18820, '/health'), (18821, '/healthcheck')):
        try:
            with urlopen(f'http://127.0.0.1:{port}{path}', timeout=2) as response:
                value = json.load(response)
                if value is True or (isinstance(value, dict) and value.get('ok') is True):
                    ready.add(port)
        except (OSError, ValueError):
            pass
    return ready


def ensure_windows_proxy():
    if os.name != 'nt':
        return
    stop = CONFIG / 'stop-proxy'
    if stop.exists():
        stop.unlink()
    missing = {18820, 18821} - reachable_ports()
    if not missing:
        return
    # A listening service may still be starting. Never take over its port.
    relay = []
    for port in sorted(missing):
        try:
            with socket.create_connection(('127.0.0.1', port), timeout=1):
                pass
        except OSError:
            relay.append(str(port))
    if relay:
        with (ROOT / 'logs' / 'windows-proxy.log').open('ab') as log:
            subprocess.Popen([sys.executable, str(ROOT / 'scripts' / 'windows_proxy.py'), '--ports', *relay],
                             stdout=log, stderr=log,
                             creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP)
    print('Waiting for local API and document editor health checks...', file=sys.stderr, flush=True)
    end = time.time() + 90
    while time.time() < end:
        missing = {18820, 18821} - reachable_ports()
        if not missing:
            return
        time.sleep(1)
    raise RuntimeError(f'Local services not ready on ports {sorted(missing)}; inspect runtime and logs/windows-proxy.log.')


def wait_saved(identifier, generation, timeout):
    end = time.time() + timeout
    while time.time() < end:
        value = api('/api/session/' + identifier + '/status')
        state = value['save']
        if state.get('state') in ('conflict', 'error'):
            raise RuntimeError('Write-back paused: ' + json.dumps(state, ensure_ascii=False))
        if state.get('state') == 'saved' and state.get('generation', 0) >= generation:
            return value
        time.sleep(0.7)
    raise TimeoutError('No confirmed disk save yet. Do not report saved.')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='action', required=True)
    for name in ('init', 'start', 'stop', 'sessions', '_watch'):
        sub.add_parser(name)
    opening = sub.add_parser('open'); opening.add_argument('file'); opening.add_argument('--new-session', action='store_true')
    for name in ('status', 'save', 'apply'):
        p = sub.add_parser(name); p.add_argument('session'); p.add_argument('--timeout', type=int, default=40)
        if name == 'apply':
            p.add_argument('--command-file', required=True); p.add_argument('--no-save', action='store_true')
    args = parser.parse_args()
    if args.action == 'init':
        return initialize()
    if args.action == '_watch':
        watch(); return {'stopped': True}
    if args.action == 'start':
        compose('up', '-d'); start_watcher(); ensure_windows_proxy(); return {'started': True, 'url': BASE}
    if args.action == 'stop':
        (CONFIG / 'stop-watcher').touch(); compose('stop'); (CONFIG / 'stop-proxy').touch(); return {'stopped': True}
    if args.action == 'sessions':
        return api('/api/sessions')
    if args.action == 'open':
        start_watcher(); return open_session(Path(args.file), args.new_session)
    if args.action == 'status':
        return api('/api/session/' + args.session + '/status')
    begin = time.time()
    if args.action == 'apply':
        body = read(Path(args.command_file))
        command = api('/api/session/' + args.session + '/commands', body)
        end = time.time() + args.timeout
        while time.time() < end:
            command = api('/api/session/' + args.session + '/commands/' + command['id'])
            if command['state'] == 'applied':
                break
            if command['state'] in ('failed', 'unknown'):
                raise RuntimeError('Live edit not confirmed: ' + json.dumps(command, ensure_ascii=False))
            time.sleep(0.6)
        else:
            raise TimeoutError('Edit result unknown. Do not resend non-idempotent commands blindly.')
        if args.no_save or body['op'] in ('inspect', 'sheet.read'):
            return command
    # Explicit force-save is the authoritative disk-write request.
    time.sleep(0.8)
    result = api('/api/session/' + args.session + '/save', {})
    if result.get('error'):
        raise RuntimeError('Document service did not accept save: ' + json.dumps(result))
    return wait_saved(args.session, result['generation'], args.timeout)


if __name__ == '__main__':
    try:
        print(json.dumps(main(), ensure_ascii=True))
    except Exception as error:
        print(f'{type(error).__name__}: {error}', file=sys.stderr)
        sys.exit(1)
