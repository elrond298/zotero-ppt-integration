"""Self-check for the local helper server.

Run it directly (no test framework needed):

    python tools/test_server.py
"""

import json
import sys
import threading
import unittest
import urllib.error
import urllib.request
from functools import partial
from http.server import ThreadingHTTPServer
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "shared"))

import zotero_proxy_server as srv  # noqa: E402


def get(port, path):
    """Returns (status, body, headers). HTTP errors are returned, not raised."""
    url = f"http://localhost:{port}{path}"
    try:
        with urllib.request.urlopen(url, timeout=5) as response:
            return response.status, response.read(), dict(response.headers)
    except urllib.error.HTTPError as e:
        return e.code, e.read(), dict(e.headers)


def post(port, path, payload):
    url = f"http://localhost:{port}{path}"
    data = json.dumps(payload).encode()
    request = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(request, timeout=5) as response:
            return response.status, response.read()
    except urllib.error.HTTPError as e:
        return e.code, e.read()


class StaticFileTest(unittest.TestCase):
    """The add-in web root must serve the pane, JavaScript, CSS and icons."""

    @classmethod
    def setUpClass(cls):
        www = srv.find_www()
        assert www is not None, "zotero-addon/www not found"
        cls.www = www
        # Plain HTTP is enough to test the handler; TLS is just a socket wrapper.
        cls.server = ThreadingHTTPServer(("localhost", 0), partial(srv.StaticHandler, directory=str(www)))
        cls.port = cls.server.server_address[1]
        threading.Thread(target=cls.server.serve_forever, daemon=True).start()

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()

    def test_root_serves_taskpane(self):
        status, body, _ = get(self.port, "/")
        self.assertEqual(status, 200)
        self.assertIn(b"Zotero Citation Manager", body)

    def test_taskpane_and_assets(self):
        for path in ("/taskpane.html", "/frontend_core.js", "/style.css", "/commands.html", "/assets/icon-32.png"):
            status, body, _ = get(self.port, path)
            self.assertEqual(status, 200, path)
            self.assertTrue(body, path)

    def test_pane_is_never_cached(self):
        _, _, headers = get(self.port, "/taskpane.html")
        self.assertEqual(headers.get("Cache-Control"), "no-store")

    def test_traversal_is_blocked(self):
        for path in ("/../server.py", "/%2e%2e/server.py", "/../shared/zotero_proxy_server.py"):
            status, _, _ = get(self.port, path)
            self.assertEqual(status, 404, path)


class ApiTest(unittest.TestCase):
    """The JSON API must answer without Zotero running."""

    @classmethod
    def setUpClass(cls):
        cls.server = srv.make_api_server(0)
        cls.port = cls.server.server_address[1]
        threading.Thread(target=cls.server.serve_forever, daemon=True).start()

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()

    def test_health(self):
        status, body, _ = get(self.port, "/health")
        self.assertEqual(status, 200)
        self.assertIn(b'"ok"', body)

    def test_bibliography_without_keys_is_rejected(self):
        status, body = post(self.port, "/bibliography", {"keys": [], "style": "apa"})
        self.assertEqual(status, 400)
        self.assertIn(b"No citation keys", body)

    def test_unknown_path_is_404(self):
        status, _, _ = get(self.port, "/nope")
        self.assertEqual(status, 404)



class DiscoveryTest(unittest.TestCase):
    def test_finds_repository_www(self):
        self.assertEqual(srv.find_www(), ROOT / "zotero-addon" / "www")

    def test_missing_cert_dir_returns_none(self):
        self.assertIsNone(srv.find_cert_dir(ROOT / "does-not-exist"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
