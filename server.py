from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "shared"))

from zotero_proxy_server import serve


if __name__ == "__main__":
    serve()
