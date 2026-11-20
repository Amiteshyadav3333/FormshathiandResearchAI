"""Native desktop shell for the local FormSaathi Flask app."""
import socket
import threading

import webview
from werkzeug.serving import make_server

from server import app


def free_port():
    with socket.socket() as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


def main():
    port = free_port()
    httpd = make_server("127.0.0.1", port, app, threaded=True)
    thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    thread.start()
    try:
        webview.create_window("FormSaathi", f"http://127.0.0.1:{port}", width=1200, height=850,
                              min_size=(900, 650), background_color="#f5f7f4")
        webview.start()
    finally:
        httpd.shutdown()


if __name__ == "__main__":
    main()
