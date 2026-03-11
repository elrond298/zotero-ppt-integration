# PowerPoint Zotero Integration

This project connects Microsoft PowerPoint to Zotero through Better BibTeX. It supports two usage modes:

- A sideloaded Office add-in in `zotero-addon/`
- A Script Lab snippet using the root `index.html`, `style.css`, and `script.js`

Core features:

- Insert in-text citations into slides
- Track citation keys in slide metadata
- View and remove stored citation keys per slide
- Generate a bibliography slide from all cited items in the presentation

## 1. Usage with npm as a sideload addon

Use this mode if you want the standalone Office add-in rather than Script Lab.

### Prerequisites

1. Install Zotero and keep it running.
2. Install Better BibTeX for Zotero.
3. Install Node.js.
4. Install Python 3. The current `npm start` flow also launches the local Python proxy.

Better BibTeX installation:

1. Download the latest `.xpi` from the Better BibTeX release page.
2. In Zotero, open `Tools > Add-ons`.
3. Use the gear menu and choose `Install Add-on From File...`.
4. Restart Zotero.

### Install and start

From the repository root:

```powershell
cd zotero-addon
npm install
npm start
```

What `npm start` does today:

- Starts the webpack dev server
- Starts the local proxy server with `uv run python server/server.py`
- Sideloads the add-in into PowerPoint desktop

If you want the debugger attached, use:

```powershell
cd zotero-addon
npm run start:debug
```

To stop the sideloaded add-in session:

```powershell
cd zotero-addon
npm run stop
```

### How to use the add-in

1. Open PowerPoint with the task pane loaded.
2. Select a slide.
3. Click `Add Citation (Pop up)` to open the Zotero picker.
4. Click `Add Citation (Selected)` to cite the currently selected Zotero item(s) directly.
5. Use the bibliography style selector to choose the output format.
6. Click `Generate Bibliography` to create a `References` slide from all stored citation keys.

### Notes

- The add-in reads and writes citation keys from slide metadata, not from slide text alone.
- If you remove citation text manually, remove the corresponding stored key from the task pane as well if you do not want it included in the bibliography.
- The task pane status indicator reflects whether the local proxy is reachable.

## 2. Usage with ScriptLab

Use this mode if you want to run the integration as a Script Lab snippet instead of a standalone add-in.

### Prerequisites

1. Install Zotero and Better BibTeX.
2. Install Python 3.
3. Install Script Lab inside PowerPoint.

Install Script Lab:

1. Open PowerPoint.
2. Go to `Insert > Get Add-ins`.
3. Search for `Script Lab` and install it.

### Start the local proxy

From the repository root:

```powershell
python server.py
```

Keep this process running while using the snippet.

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
- `zotero-addon/` contains the standalone Office add-in.
- `shared/frontend_core.js` is the single source of truth for frontend behavior.
- `shared/zotero_proxy_server.py` is the single source of truth for the local proxy behavior.

### Generated and wrapped files

These files should not be treated as primary edit targets:

- `script.js`
- `zotero-addon/src/taskpane/taskpane.ts`
- `server.py`
- `zotero-addon/server/server.py`

For frontend behavior changes, edit `shared/frontend_core.js` and regenerate the outputs.

For backend proxy changes, edit `shared/zotero_proxy_server.py`.

### Regenerating shared frontend files

From the repository root:

```powershell
python.exe tools/sync_shared.py
```

This updates:

- `script.js`
- `zotero-addon/src/taskpane/taskpane.ts`

### Validation

Python validation:

```powershell
python -m py_compile server.py zotero-addon/server/server.py shared/zotero_proxy_server.py tools/sync_shared.py
```

Add-in build validation:

```powershell
cd zotero-addon
npm run build
```

Manifest validation:

```powershell
cd zotero-addon
npm run validate
```

### Implementation notes

- The frontend calls the local proxy on `http://localhost:8000`.
- `/zotero` proxies Better BibTeX CAYW calls.
- `/bibliography` proxies Better BibTeX JSON-RPC bibliography generation.
- `/health` is used by the UI status indicator.
- Bibliography output is based on citation keys stored in slide metadata.

### Useful links

- https://retorque.re/zotero-better-bibtex/citing/cayw
- https://retorque.re/zotero-better-bibtex/exporting/json-rpc/index.html
- https://www.zotero.org/support/dev/web_api/v3/basics
- https://learn.microsoft.com/en-us/javascript/api/powerpoint
