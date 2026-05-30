import errno
import os
import socket
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.request

from werkzeug.serving import make_server

from index import app

HOST = "127.0.0.1"
PORT_FILE = ".swim_final_input_app_port"
PORT_START = 5155
PORT_END = 5199


def app_url(port):
    return f"http://{HOST}:{port}"


def health_url(port):
    return f"{app_url(port)}/health"


def read_saved_port():
    try:
        with open(PORT_FILE, "r", encoding="utf-8") as file:
            return int(file.read().strip())
    except (OSError, ValueError):
        return None


def write_saved_port(port):
    with open(PORT_FILE, "w", encoding="utf-8") as file:
        file.write(str(port))


def url_contains_app(port):
    try:
        with urllib.request.urlopen(health_url(port), timeout=0.4) as response:
            if response.read().decode("utf-8") == "swim-data-app-final-input-ok":
                return True
    except (OSError, urllib.error.URLError):
        pass

    try:
        with urllib.request.urlopen(app_url(port), timeout=0.4) as response:
            return "水泳データ集計" in response.read().decode("utf-8", errors="ignore")
    except (OSError, urllib.error.URLError):
        return False


def find_running_app():
    saved_port = read_saved_port()
    if saved_port and url_contains_app(saved_port):
        return saved_port
    for port in range(PORT_START, PORT_END + 1):
        if url_contains_app(port):
            write_saved_port(port)
            return port
    return None


def open_browser(port):
    url = app_url(port)
    if os.environ.get("SWIM_APP_NO_BROWSER") == "1":
        print(f"browser skipped: {url}", flush=True)
        return
    if sys.platform == "darwin":
        subprocess.run(["/usr/bin/open", url], check=False)
    else:
        import webbrowser
        webbrowser.open(url)


def find_free_port(start=PORT_START, end=PORT_END):
    for port in range(start, end + 1):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                sock.bind((HOST, port))
            except OSError as exc:
                if exc.errno == errno.EADDRINUSE:
                    continue
                raise
            return port
    raise RuntimeError("空きポートが見つかりませんでした")


class LocalServer(threading.Thread):
    def __init__(self, host, port):
        super().__init__(daemon=True)
        self.server = make_server(host, port, app)
        self.context = app.app_context()
        self.context.push()

    def run(self):
        self.server.serve_forever()

    def shutdown(self):
        self.server.shutdown()


def main():
    running_port = find_running_app()
    if running_port:
        print(f"using existing server: {app_url(running_port)}", flush=True)
        open_browser(running_port)
        return

    port = find_free_port()
    write_saved_port(port)
    server = LocalServer(HOST, port)
    server.start()
    print(f"started server: {app_url(port)}", flush=True)

    deadline = time.time() + 5
    while time.time() < deadline:
        if url_contains_app(port):
            break
        time.sleep(0.1)
    open_browser(port)

    test_seconds = os.environ.get("SWIM_APP_TEST_SECONDS")
    if test_seconds:
        time.sleep(float(test_seconds))
        server.shutdown()
        return

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        server.shutdown()


if __name__ == "__main__":
    main()
