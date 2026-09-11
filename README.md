# PowerPoint Zotero Integration

This project connects Microsoft PowerPoint to Zotero through Better BibTeX. It supports two usage modes:

- A sideloaded Office add-in in `zotero-addon/` (installs without Node.js or npm)
- A Script Lab snippet using the root `index.html`, `style.css`, and `script.js`

Core features:

- Insert in-text citations into slides
- Track citation keys in slide metadata
- View and remove stored citation keys per slide
- Generate a bibliography slide from all cited items in the presentation

## 1. Install the add-in

### Prerequisites

1. Install Zotero and keep it running.
2. Install Better BibTeX for Zotero.
3. Install Python 3 — either from [python.org](https://www.python.org/downloads/) or with `winget install astral-sh.uv`.
4. Have PowerPoint installed on Windows.

Node.js, npm and a build step are **not** required: the add-in is plain HTML/JavaScript, and the
helper server below uses only the Python standard library.

Better BibTeX installation:

1. Download the latest `.xpi` from the Better BibTeX release page.
2. In Zotero, open `Tools > Add-ons`.
3. Use the gear menu and choose `Install Add-on From File...`.
4. Restart Zotero.

### Install

From the repository root, in Windows PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

The same command works from WSL, because Windows PowerShell can read the repository over the
`\\wsl.localhost` path:

```bash
powershell.exe -NoProfile -ExecutionPolicy Bypass -File '\\wsl.localhost\<distro>\<path-to-repo>\install.ps1'
```

`install.ps1` does all of this, and can be re-run at any time to update the installed copy:

- copies the add-in web files and the helper server into `%LOCALAPPDATA%\ZoteroCitations`
- creates a trusted localhost certificate for `https://localhost:23000` (reused if it is still valid)
- registers the add-in with PowerPoint (`HKCU:\Software\Microsoft\Office\16.0\Wef\Developer`)
- adds a Start-up shortcut so the helper server runs at every logon and starts it right away
- checks that `https://localhost:23000/taskpane.html` answers, so failures are visible immediately

Use `-NoAutostart` if you prefer to start the server by hand with
`%LOCALAPPDATA%\ZoteroCitations\run-server.cmd`.

### How to use the add-in

1. Restart PowerPoint after installing (Office reads the add-in registration at start-up).
2. On the Home tab, open `Zotero Tools > Open Zotero Pane`.
3. Select a slide.
4. Click `Add Citation (Pop up)` to open the Zotero picker.
5. Click `Add Citation (Selected)` to cite the currently selected Zotero item(s) directly.
6. Use the bibliography style selector to choose the output format.
7. Click `Generate Bibliography` to create a `References` slide from all stored citation keys.

### Notes

- The add-in reads and writes citation keys from slide metadata, not from slide text alone.
- If you remove citation text manually, remove the corresponding stored key from the task pane as well if you do not want it included in the bibliography.
- The task pane status indicator reflects whether the local helper server is reachable.

### Troubleshooting

- **The pane stays blank.** Read `%LOCALAPPDATA%\ZoteroCitations\server.log`, and open
  `https://localhost:23000/taskpane.html` in a browser to see what the add-in receives.
- **"Proxy not running" in the pane.** The helper server is not running: start
  `%LOCALAPPDATA%\ZoteroCitations\run-server.cmd` (visible console, useful for logs) or re-run `install.ps1`.
- **`FATAL: Could not bind to port ...`** means another copy of the server is already running.
- **The pane stops loading after about a year.** The localhost certificate expired: re-run `install.ps1` to renew it.
- **The pane cannot reach anything at all** on some Office builds. Enable loopback for the Office
  webview once, in an **elevated** PowerShell, and restart PowerPoint:

  ```powershell
  npx office-addin-dev-settings appcontainer EdgeWebView --loopback
  ```

- **Uninstall:**

  ```powershell
  powershell -ExecutionPolicy Bypass -File .\uninstall.ps1
  ```

## 2. Usage with ScriptLab

Use this mode if you want to run the integration as a Script Lab snippet instead of the add-in.

### Prerequisites

1. Install Zotero and Better BibTeX.
2. Install Python 3.
3. Install Script Lab inside PowerPoint (`Insert > Get Add-ins`, search for `Script Lab`).

### Start the local proxy

The helper server installed in step 1 already listens on `http://localhost:8000`, so normally there is
nothing to start. Without the add-in installed, run one of these from the repository root and keep it
running:

```powershell
python server.py
```

```bash
uv run python server.py
```

The server serves the JSON API on `http://localhost:8000` and, when a localhost certificate is present,
also the add-in files on `https://localhost:23000`. Use `--no-static` to serve only the API.

### Load the Script Lab snippet

1. Open Script Lab in PowerPoint.
2. Create a new snippet.
3. Copy these files into the matching tabs:

- `index.html` -> HTML
- `style.css` -> CSS
- `script.js` -> Script

No extra libraries are required.

### How to use the snippet

1. Run the snippet.
2. Select a slide.
3. Use `Add Citation (Pop up)` or `Add Citation (Selected)`.
4. Choose a bibliography style from the selector.
5. Click `Generate Bibliography` when you are ready to build the references slide.

### Notes

- Script Lab uses the same frontend logic as the standalone add-in.
- The local proxy is still required because the Office webview cannot reliably talk directly to Better BibTeX on `127.0.0.1:23119`.

## 3. Development notes

### Repository structure

- Root files `index.html`, `style.css`, and `script.js` are for Script Lab usage.
- `zotero-addon/www/` is the add-in web root: `taskpane.html`, `commands.html`, `style.css`, `frontend_core.js`.
- `zotero-addon/manifest.xml` is the add-in manifest; `install.ps1` copies it to `%LOCALAPPDATA%\ZoteroCitations`.
- `shared/frontend_core.js` is the single source of truth for frontend behavior.
- `shared/zotero_proxy_server.py` is the single source of truth for the helper server.
- `install.ps1` / `uninstall.ps1` install and remove the add-in on Windows.

### Generated and wrapped files

These files should not be treated as primary edit targets:

- `script.js`
- `zotero-addon/www/frontend_core.js`
- `zotero-addon/www/style.css`
- `server.py`

For frontend behavior changes, edit `shared/frontend_core.js` and regenerate the outputs.
For style changes, edit `style.css` and regenerate the outputs.
For backend changes, edit `shared/zotero_proxy_server.py`; `install.ps1` copies it to the install folder.

### Regenerating shared frontend files

```bash
python tools/sync_shared.py
```

This updates `script.js`, `zotero-addon/www/frontend_core.js`, and `zotero-addon/www/style.css`.
Afterwards re-run `install.ps1` so PowerPoint gets the new files (they are served from the install folder).

### Validation

```bash
python -m py_compile server.py shared/zotero_proxy_server.py tools/sync_shared.py
python tools/test_server.py
```

### Implementation notes

- The add-in fetches the helper server on `http://localhost:8000`.
- `/zotero` proxies Better BibTeX CAYW calls.
- `/bibliography` proxies Better BibTeX JSON-RPC bibliography generation.
- `/health` is used by the UI status indicator.
- `https://localhost:23000` serves the pane itself (Office requires HTTPS for task panes).
- Bibliography output is based on citation keys stored in slide metadata.

### Useful links

- https://retorque.re/zotero-better-bibtex/citing/cayw
- https://retorque.re/zotero-better-bibtex/exporting/json-rpc/index.html
- https://www.zotero.org/support/dev/web_api/v3/basics
- https://learn.microsoft.com/en-us/javascript/api/powerpoint
