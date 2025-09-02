from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.request import urlopen, Request
from urllib.error import URLError
import json
import traceback

PORT = 8000
# Endpoint for the initial citation picker
ZOTERO_CAYW_ENDPOINT = "http://127.0.0.1:23119/better-bibtex/cayw?format=json"
# Endpoint for the new bibliography generation
BBT_JSONRPC_ENDPOINT = "http://127.0.0.1:23119/better-bibtex/json-rpc"

class ProxyHandler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        """Handle pre-flight CORS requests."""
        self.send_response(200, "ok")
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header("Access-Control-Allow-Headers", "X-Requested-With, Content-Type")
        self.end_headers()

    def do_GET(self):
        if self.path == '/zotero':
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
                # 1. Read and parse the request body from the PowerPoint add-in
                content_length = int(self.headers['Content-Length'])
                post_data = self.rfile.read(content_length)
                request_body = json.loads(post_data)
                
                keys = request_body.get('keys', [])
                # Note: This style must be a valid CSL style ID (e.g., 'apa', 'chicago-author-date')
                # 'apalike' is a BibTeX style, so we'll map it to 'apa' which is the CSL equivalent.
                style_name = request_body.get('style', 'apa')
                if style_name == 'apalike':
                    style_name = 'apa' # Map BibTeX style to CSL style

                if not keys:
                    self._send_json_response(400, {"error": "No citation keys provided"})
                    return

                # 2. Construct the JSON-RPC payload for the BBT server
                payload = {
                    "jsonrpc": "2.0",
                    "method": "item.bibliography",
                    "params": [
                        keys,  # First parameter: array of citekeys
                        {      # Second parameter: format object
                            "id": style_name,
                            "contentType": "text" # We want plain text for the slide
                        }
                    ]
                }
                
                # 3. Send the request to the BBT JSON-RPC endpoint
                req_data = json.dumps(payload).encode('utf-8')
                req = Request(BBT_JSONRPC_ENDPOINT, data=req_data, headers={
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                })

                with urlopen(req) as response:
                    response_data = json.loads(response.read())

                # 4. Process the response from BBT
                if 'error' in response_data:
                    # Forward the error from BBT to the client
                    error_info = response_data['error']
                    print(f"BBT JSON-RPC Error: {error_info.get('message')}")
                    self._send_json_response(500, {"error": f"Zotero/BBT Error: {error_info.get('message')}"})
                else:
                    # Extract the successful result
                    bibliography_text = response_data.get('result', '')
                    print(bibliography_text)
                    self._send_json_response(200, {"bibliography": bibliography_text})

            except URLError as e:
                error_message = "Could not connect to Zotero/BBT. Is Zotero running with the Better BibTeX plugin installed?"
                print(f"Error connecting to BBT JSON-RPC endpoint: {e}")
                self._send_json_response(500, {"error": error_message})
            except Exception as e:
                print("An unexpected error occurred during bibliography generation:")
                print(traceback.format_exc())
                self._send_json_response(500, {"error": f"An internal server error occurred: {e}"})
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
    server = HTTPServer(('localhost', PORT), ProxyHandler)
    print(f"Proxy server running on http://localhost:{PORT}")
    print("Forwarding bibliography requests to Zotero/Better BibTeX JSON-RPC endpoint.")
    server.serve_forever()
