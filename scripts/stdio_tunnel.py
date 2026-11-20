"""Relay one stdio connection to an editor port on WSL loopback."""
import socket
import sys
import threading


def main():
    port = int(sys.argv[1])
    if port not in (18820, 18821):
        raise ValueError('Editor ports only')
    connection = socket.create_connection(('127.0.0.1', port), timeout=15)
    connection.settimeout(None)

    def upload():
        try:
            while chunk := sys.stdin.buffer.read1(65536):
                connection.sendall(chunk)
        finally:
            try:
                connection.shutdown(socket.SHUT_WR)
            except OSError:
                pass

    threading.Thread(target=upload, daemon=True).start()
    try:
        while chunk := connection.recv(65536):
            sys.stdout.buffer.write(chunk)
            sys.stdout.buffer.flush()
    finally:
        connection.close()


if __name__ == '__main__':
    main()
