from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.request import urlopen
from urllib.error import URLError
import json

PORT = 8000
ZOTERO_ENDPOINT = "http://127.0.0.1:23119/better-bibtex/cayw?format=json"

class ProxyHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path == '/zotero':
            try:
                with urlopen(ZOTERO_ENDPOINT) as response:
                    data = response.read()
                    self.send_response(200)
                    self.send_header('Content-Type', 'application/json')
                    self.send_header('Access-Control-Allow-Origin', '*')
                    self.end_headers()
                    self.wfile.write(data)
            except URLError as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}).encode())
        else:
            self.send_response(404)
            self.end_headers()

if __name__ == '__main__':
    server = HTTPServer(('localhost', PORT), ProxyHandler)
    print(f"Proxy server running on http://localhost:{PORT}")
    server.serve_forever()

