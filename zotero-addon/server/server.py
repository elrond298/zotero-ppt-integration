# Copy of root-level server with minor path isolation.
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.request import urlopen, Request
from urllib.error import URLError
import json
import traceback
import time
import sys
import socket

# Simple timestamp helper
def _ts():
    return time.strftime('%Y-%m-%d %H:%M:%S')

PORT = 8000
ZOTERO_CAYW_ENDPOINT = "http://127.0.0.1:23119/better-bibtex/cayw?format=json"
BBT_JSONRPC_ENDPOINT = "http://127.0.0.1:23119/better-bibtex/json-rpc"

class ProxyHandler(BaseHTTPRequestHandler):
    # More compact structured log output
    def log_message(self, format, *args):  # noqa: D401 (override)
        # Default implementation writes to stderr; keep that but add timestamp prefix
        sys.stderr.write(f"[{_ts()}] {self.address_string()} - {format % args}\n")

    def do_OPTIONS(self):
        self.send_response(200, "ok")
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header("Access-Control-Allow-Headers", "X-Requested-With, Content-Type")
        self.end_headers()

    def do_GET(self):
        if self.path == '/health':
            self._send_json_response(200, {"status": "ok"})
        elif self.path == '/zotero':
            try:
                with urlopen(ZOTERO_CAYW_ENDPOINT) as response:
                    data = response.read()
                    self.send_response(200)
                    self.send_header('Content-Type', 'application/json')
                    self.send_header('Access-Control-Allow-Origin', '*')
                    self.end_headers()
                    self.wfile.write(data)
            except URLError as e:
                error_message = "Could not connect to Zotero/BBT. Is Zotero running with the Better BibTeX plugin installed?"
                print(f"Error connecting to Zotero CAYW endpoint: {e}")
                self._send_json_response(500, {"error": error_message})
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        if self.path == '/bibliography':
            try:
                content_length = int(self.headers['Content-Length'])
                post_data = self.rfile.read(content_length)
                request_body = json.loads(post_data)
                keys = request_body.get('keys', [])
                style_name = request_body.get('style', 'apa')
                if style_name == 'apalike':
                    style_name = 'apa'
                if not keys:
                    self._send_json_response(400, {"error": "No citation keys provided"})
                    return
                payload = {
                    "jsonrpc": "2.0",
                    "method": "item.bibliography",
                    "params": [
                        keys,
                        {"id": style_name, "contentType": "text"}
                    ]
                }
                req_data = json.dumps(payload).encode('utf-8')
                req = Request(BBT_JSONRPC_ENDPOINT, data=req_data, headers={'Content-Type': 'application/json','Accept': 'application/json'})
                with urlopen(req) as response:
                    response_data = json.loads(response.read())
                if 'error' in response_data:
                    error_info = response_data['error']
                    print(f"BBT JSON-RPC Error: {error_info.get('message')}")
                    self._send_json_response(500, {"error": f"Zotero/BBT Error: {error_info.get('message')}"})
                else:
                    bibliography_text = response_data.get('result', '')
                    print(bibliography_text)
                    self._send_json_response(200, {"bibliography": bibliography_text})
            except URLError as e:
                error_message = "Could not connect to Zotero/BBT. Is Zotero running with the Better BibTeX plugin installed?"
                print(f"Error connecting to BBT JSON-RPC endpoint: {e}")
                self._send_json_response(500, {"error": error_message})
            except Exception as e:
                print("Unexpected error during bibliography generation:")
                print(traceback.format_exc())
                self._send_json_response(500, {"error": f"Internal server error: {e}"})
        else:
            self.send_response(404)
            self.end_headers()

    def _send_json_response(self, status_code, data):
        self.send_response(status_code)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(json.dumps(data).encode())

if __name__ == '__main__':
    try:
        server = HTTPServer(('localhost', PORT), ProxyHandler)
    except OSError as e:
        # Common cause: port already in use -> leads to exit code 1 without clear context
        print(f"[{_ts()}] FATAL: Could not bind to port {PORT} (likely already in use). Error: {e}", flush=True)
        # Give user a quick hint on what holds the port (Windows netstat example)
        if hasattr(e, 'winerror') and e.winerror == 10048:  # Address in use
            print("Hint: Another instance may be running. Close it or free the port. Example check: netstat -ano | findstr :8000", flush=True)
        sys.exit(1)
    print(f"[{_ts()}] Local proxy server running on http://localhost:{PORT}", flush=True)
    print(f"[{_ts()}] Forwarding /zotero -> {ZOTERO_CAYW_ENDPOINT}", flush=True)
    print(f"[{_ts()}] Forwarding /bibliography -> {BBT_JSONRPC_ENDPOINT}", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print(f"[{_ts()}] Shutdown requested (Ctrl+C)", flush=True)
    except Exception as e:  # Catch unexpected to avoid silent exit code 1
        print(f"[{_ts()}] FATAL: Unhandled exception in serve_forever: {e}", flush=True)
        print(traceback.format_exc())
        sys.exit(1)
    finally:
        try:
            server.server_close()
        except Exception:
            pass
        print(f"[{_ts()}] Server stopped", flush=True)
