"""Relay selected Windows loopback ports without changing WSL or firewall settings."""
import argparse
import json
from pathlib import Path
import socketserver
import subprocess
import threading
import time

ROOT = Path(__file__).resolve().parent.parent
CONFIG = ROOT / 'config'


class Relay(socketserver.BaseRequestHandler):
    def handle(self):
        cfg = json.loads((CONFIG / 'settings.json').read_text(encoding='utf-8'))
        source = ROOT / 'scripts' / 'stdio_tunnel.py'
        linux = '/mnt/' + source.drive[0].lower() + '/' + str(source)[3:].replace('\\', '/')
        process = subprocess.Popen(['wsl', '-d', cfg['distro'], '-u', 'root', '--',
                                    'python3', linux, str(self.server.server_address[1])],
                                   stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
                                   creationflags=subprocess.CREATE_NO_WINDOW)

        def upload():
            try:
                while chunk := self.request.recv(65536):
                    process.stdin.write(chunk)
                    process.stdin.flush()
            except (OSError, BrokenPipeError):
                pass
            finally:
                try:
                    process.stdin.close()
                except OSError:
                    pass

        threading.Thread(target=upload, daemon=True).start()
        try:
            while chunk := process.stdout.read1(65536):
                self.request.sendall(chunk)
        except (OSError, BrokenPipeError):
            pass
        finally:
            self.request.close()
            # Only the relay subprocess created for this socket is stopped.
            if process.poll() is None:
                process.terminate()
            process.wait(timeout=5)


class Server(socketserver.ThreadingTCPServer):
    allow_reuse_address = True
    daemon_threads = True


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--ports', type=int, nargs='+', choices=(18820, 18821), required=True)
    args = parser.parse_args()
    cfg = json.loads((CONFIG / 'settings.json').read_text(encoding='utf-8'))
    keepalive = subprocess.Popen(['wsl', '-d', cfg['distro'], '-u', 'root', '--', 'sleep', 'infinity'],
                                 stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                                 stderr=subprocess.DEVNULL, creationflags=subprocess.CREATE_NO_WINDOW)
    servers = []
    try:
        for port in sorted(set(args.ports)):
            server = Server(('127.0.0.1', port), Relay)
            servers.append(server)
            threading.Thread(target=server.serve_forever, daemon=True).start()
        while not (CONFIG / 'stop-proxy').exists():
            time.sleep(1)
    finally:
        for server in servers:
            server.shutdown()
            server.server_close()
        keepalive.terminate()
        keepalive.wait(timeout=5)
