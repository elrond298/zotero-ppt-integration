"""Local helper server for the Zotero -> PowerPoint integration.

One process serves both halves of the add-in:

* ``http://localhost:8000``  -- JSON API (used by the Script Lab snippet and the add-in)
* ``https://localhost:23000`` -- the add-in web files (Office requires HTTPS for task panes)

Only the Python standard library is required.
"""

from __future__ import annotations

import argparse
import json
import os
import ssl
import sys
import threading
import time
import traceback
from functools import partial
from http.server import BaseHTTPRequestHandler, SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.error import URLError
from urllib.request import Request, urlopen

HERE = Path(__file__).resolve().parent

API_PORT = 8000
STATIC_PORT = 23000

ZOTERO_CAYW_ENDPOINT = "http://127.0.0.1:23119/better-bibtex/cayw?format=json"
BBT_JSONRPC_ENDPOINT = "http://127.0.0.1:23119/better-bibtex/json-rpc"

CERT_DIR_NAME = ".office-addin-dev-certs"
CERT_ENV_VAR = "ZOTERO_PPT_CERT_DIR"

# The add-in web files sit next to this file once installed, and in
# zotero-addon/www when the server runs straight from the repository.
WWW_CANDIDATES = (HERE / "www", HERE.parent / "zotero-addon" / "www")

ZOTERO_DOWN_ERROR = (
    "Could not connect to Zotero/BBT. Is Zotero running with the Better BibTeX plugin installed?"
)


def _ts():
    return time.strftime("%Y-%m-%d %H:%M:%S")


class ProxyHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        sys.stderr.write(f"[{_ts()}] {self.address_string()} - {format % args}\n")

    def do_OPTIONS(self):
        self.send_response(200, "ok")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "X-Requested-With, Content-Type")
        self.end_headers()

    def do_GET(self):
        if self.path == "/health":
            self._send_json_response(200, {"status": "ok"})
        elif self.path.startswith("/zotero"):
            try:
                selected = "selected=true" in self.path
                endpoint = ZOTERO_CAYW_ENDPOINT
                if selected:
                    endpoint += "&selected=true"

                with urlopen(endpoint) as response:
                    data = response.read()
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(data)
            except URLError as e:
                print(f"Error connecting to Zotero CAYW endpoint: {e}")
                self._send_json_response(500, {"error": ZOTERO_DOWN_ERROR})
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        if self.path == "/bibliography":
            try:
                content_length = int(self.headers["Content-Length"])
                post_data = self.rfile.read(content_length)
                request_body = json.loads(post_data)
                keys = request_body.get("keys", [])
                style_name = request_body.get("style", "apa")
                if style_name == "apalike":
                    style_name = "apa"
                if not keys:
                    self._send_json_response(400, {"error": "No citation keys provided"})
                    return

                payload = {
                    "jsonrpc": "2.0",
                    "method": "item.bibliography",
                    "params": [
                        keys,
                        {"id": style_name, "contentType": "text"},
                    ],
                }
                req_data = json.dumps(payload).encode("utf-8")
                req = Request(
                    BBT_JSONRPC_ENDPOINT,
                    data=req_data,
                    headers={"Content-Type": "application/json", "Accept": "application/json"},
                )
                with urlopen(req) as response:
                    response_data = json.loads(response.read())
                if "error" in response_data:
                    error_info = response_data["error"]
                    print(f"BBT JSON-RPC Error: {error_info.get('message')}")
                    self._send_json_response(500, {"error": f"Zotero/BBT Error: {error_info.get('message')}"})
                else:
                    bibliography_text = response_data.get("result", "")
                    print(bibliography_text)
                    self._send_json_response(200, {"bibliography": bibliography_text})
            except URLError as e:
                print(f"Error connecting to BBT JSON-RPC endpoint: {e}")
                self._send_json_response(500, {"error": ZOTERO_DOWN_ERROR})
            except Exception as e:
                print("Unexpected error during bibliography generation:")
                print(traceback.format_exc())
                self._send_json_response(500, {"error": f"Internal server error: {e}"})
        else:
            self.send_response(404)
            self.end_headers()

    def _send_json_response(self, status_code, data):
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(json.dumps(data).encode())


class StaticHandler(SimpleHTTPRequestHandler):
    """Serves the add-in web files. ``/`` shows the task pane."""

    def do_GET(self):
        if self.path in ("", "/"):
            self.path = "/taskpane.html"
        super().do_GET()

    def end_headers(self):
        # The pane is a dev install: never let WebView2 cache stale files.
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def log_message(self, format, *args):
        sys.stderr.write(f"[{_ts()}] static {self.address_string()} - {format % args}\n")


def find_www(explicit=None):
    if explicit:
        path = Path(explicit).expanduser()
        return path if path.is_dir() else None
    for candidate in WWW_CANDIDATES:
        if candidate.is_dir():
            return candidate
    return None


def find_cert_dir(explicit=None):
    """Locate a directory holding localhost.crt/localhost.key, else None."""
    for candidate in _cert_dir_candidates(explicit):
        if (candidate / "localhost.crt").is_file() and (candidate / "localhost.key").is_file():
            return candidate
    return None


def _cert_dir_candidates(explicit):
    if explicit:
        # An explicit --cert-dir is authoritative: never silently fall back to another one.
        yield Path(explicit).expanduser()
        return
    env_dir = os.environ.get(CERT_ENV_VAR)
    if env_dir:
        yield Path(env_dir).expanduser()
    yield Path.home() / CERT_DIR_NAME
    windows_users = Path("/mnt/c/Users")
    if windows_users.is_dir():  # WSL: reuse the certificate the Windows install created
        for user_dir in sorted(windows_users.glob("*")):
            yield user_dir / CERT_DIR_NAME


def _exit_port_in_use(port, e):
    print(f"[{_ts()}] FATAL: Could not bind to port {port} (likely already in use). Error: {e}", flush=True)
    if getattr(e, "winerror", None) == 10048:
        print("Hint: the helper server installed for the add-in may already be running.", flush=True)
        print("      It already serves http://localhost:8000, so the Script Lab snippet works without starting anything.", flush=True)
    sys.exit(1)


def make_api_server(port=API_PORT):
    try:
        return ThreadingHTTPServer(("localhost", port), ProxyHandler)
    except OSError as e:
        _exit_port_in_use(port, e)


def make_static_server(port, www_dir, cert_dir):
    cert_file = Path(cert_dir) / "localhost.crt"
    key_file = Path(cert_dir) / "localhost.key"
    try:
        server = ThreadingHTTPServer(("localhost", port), partial(StaticHandler, directory=str(www_dir)))
    except OSError as e:
        _exit_port_in_use(port, e)

    context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.load_cert_chain(certfile=str(cert_file), keyfile=str(key_file))
    server.socket = context.wrap_socket(server.socket, server_side=True)
    return server


def _redirect_output(path):
    handle = open(path, "a", encoding="utf-8", buffering=1)
    sys.stdout = handle
    sys.stderr = handle


def serve(api_port=API_PORT, static_port=STATIC_PORT, with_static=True, cert_dir=None, www_dir=None, log_path=None):
    if log_path:
        _redirect_output(log_path)

    servers = [make_api_server(api_port)]
    print(f"[{_ts()}] Local proxy server running on http://localhost:{api_port}", flush=True)
    print(f"[{_ts()}] Forwarding /zotero -> {ZOTERO_CAYW_ENDPOINT}", flush=True)
    print(f"[{_ts()}] Forwarding /bibliography -> {BBT_JSONRPC_ENDPOINT}", flush=True)

    if with_static:
        www = find_www(www_dir)
        certs = find_cert_dir(cert_dir)
        if www is None:
            print(f"[{_ts()}] Add-in web files not found (looked in {', '.join(str(c) for c in WWW_CANDIDATES)}).", flush=True)
            print("[{0}] The add-in task pane will not load; the Script Lab snippet still works.".format(_ts()), flush=True)
        elif certs is None:
            print(f"[{_ts()}] No localhost certificate found. Run install.ps1 to create one.", flush=True)
            print(f"[{_ts()}] The add-in task pane will not load; the Script Lab snippet still works.", flush=True)
        else:
            server = make_static_server(static_port, www, certs)
            servers.append(server)
            print(f"[{_ts()}] Add-in files served on https://localhost:{static_port} from {www}", flush=True)

    for server in servers:
        threading.Thread(target=server.serve_forever, daemon=True).start()

    try:
        while True:
            time.sleep(3600)
    except KeyboardInterrupt:
        print(f"[{_ts()}] Shutdown requested (Ctrl+C)", flush=True)
    finally:
        for server in servers:
            server.shutdown()
            server.server_close()
        print(f"[{_ts()}] Server stopped", flush=True)


def build_parser():
    parser = argparse.ArgumentParser(description="Zotero <-> PowerPoint helper server")
    parser.add_argument("--api-port", type=int, default=API_PORT, help=f"JSON API port (default {API_PORT})")
    parser.add_argument(
        "--static-port", type=int, default=STATIC_PORT, help=f"HTTPS add-in port (default {STATIC_PORT})"
    )
    parser.add_argument("--no-static", action="store_true", help="serve only the JSON API (Script Lab mode)")
    parser.add_argument("--www", help="directory holding the add-in web files")
    parser.add_argument("--cert-dir", help=f"directory holding localhost.crt/localhost.key (or ${CERT_ENV_VAR})")
    parser.add_argument("--log", help="append all output to this file (used for the windowless autostart)")
    return parser


def main(argv=None):
    args = build_parser().parse_args(argv)
    serve(
        api_port=args.api_port,
        static_port=args.static_port,
        with_static=not args.no_static,
        cert_dir=args.cert_dir,
        www_dir=args.www,
        log_path=args.log,
    )


if __name__ == "__main__":
    main()
