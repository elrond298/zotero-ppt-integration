# PowerPoint Zotero Integration

This project connects Microsoft PowerPoint to Zotero through Better BibTeX. It supports two usage modes:

- A sideloaded Office add-in in `zotero-addon/` (installs without Node.js, npm or Python)
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
3. Have PowerPoint installed on Windows.

Nothing else: the local helper is a small C# program that `install.ps1` compiles with the C# compiler
that ships with Windows (`csc.exe`), and it uses only the .NET Framework. No Node.js, no npm, no
Python, no `pip`, no downloads.

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

- compiles `zotero-addon/helper/ZoteroHelper.cs` and installs it with the web files and the manifest
  in `%LOCALAPPDATA%\ZoteroCitations`
- creates a trusted localhost certificate for `https://localhost:23000` (reused while it is valid)
- registers the add-in with PowerPoint (`HKCU:\Software\Microsoft\Office\16.0\Wef\Developer`)
- adds a Start-up shortcut (a hidden VBScript launcher) so the helper runs at every logon, and starts it now
- checks that `https://localhost:23000/taskpane.html` answers, so failures are visible immediately

Use `-NoAutostart` if you prefer to start the helper by hand with
`%LOCALAPPDATA%\ZoteroCitations\run-server.cmd` (that one keeps a console window open, which is handy
for reading the log).

### How to use the add-in

1. Restart PowerPoint after installing (Office reads the add-in registration at start-up).
2. Start it from the Add-ins tab: `Home > Add-ins (加载项) > Zotero Citations`. The `Zotero Tools`
   group (with `Open Zotero Pane`) appears on the Home tab while the add-in is loaded, but Office does
   not pin sideloaded add-ins there permanently (see Troubleshooting).
3. Select a slide.
4. Click `Add Citation (Pop up)` to open the Zotero picker. Anything you fill in there (`see `, `p. 42`,
   suppress author) is written into the citation text.
5. Click `Add Citation (Selected)` to cite the currently selected Zotero item(s) directly.
6. Use the bibliography style selector to choose the output format.
7. Click `Generate Bibliography`. It refreshes the deck's References slide (or creates one) from all
   stored citation keys, with the formatting the chosen style asks for (italics, bold, sub- and
   superscript). Entries that do not fit continue on `References (cont.)` slides while the checkbox
   under the style selector is ticked (the default); untick it to always use a single slide.

### Notes

- The add-in reads and writes citation keys from slide metadata, not from slide text alone.
- Slide reordering uses `Slide.moveTo`, which some PowerPoint builds (16.0.17932 on Windows, for one)
  reject with a GeneralException; when that happens the reordering is skipped and the bibliography
  slides stay where they are, with a note in `server.log`.
- The slides we generate are tagged (`ZOTERO_BIBLIOGRAPHY`, with the page number as the value), so
  generating again rewrites those slides instead of adding more, a slide titled `References` from an
  older version is adopted, continuation slides that are no longer needed are deleted, and the whole
  block is kept at the end of the deck.
- Citation keys that Zotero cannot resolve (renamed or deleted items) are listed in the pane instead of
  being silently dropped from the bibliography.
- If you remove citation text manually, remove the corresponding stored key from the task pane as well if you do not want it included in the bibliography.
- The task pane status indicator reflects whether the local helper is reachable.
- The helper is two listeners in one process: the JSON API on `http://localhost:8000` and the pane
  itself on `https://localhost:23000`. Office requires HTTPS for task panes, and Office's webview
  cannot reach Better BibTeX directly.

### Troubleshooting

- A pane that loaded while the helper was restarting reloads itself once to get its stylesheet back.
- **The pane stays blank or looks unstyled.** Read `%LOCALAPPDATA%\ZoteroCitations\server.log`, and open
  `https://localhost:23000/taskpane.html` in a browser to see what the add-in receives.
- **"Proxy not running" in the pane.** The helper is not running: start
  `%LOCALAPPDATA%\ZoteroCitations\run-server.cmd` (visible console with the same log) or re-run `install.ps1`.
- **`FATAL: could not bind to port ...`** means another copy of the helper is already running.
- **The pane stops loading after about a year.** The localhost certificate expired: re-run `install.ps1` to renew it.
- **The add-in is not pinned to the Home tab.** Office keeps sideloaded (developer) add-ins in
  `Home > Add-ins` (加载项); starting it there opens the pane and adds the `Zotero Tools` group for that
  session. A permanently pinned button needs Microsoft's deployment path (AppSource or the Microsoft 365
  admin center), which requires the add-in files to be hosted on a public HTTPS URL - and a publicly
  hosted pane could not reach this local helper anyway (WebView2 blocks public pages from calling
  localhost). The deployment options were checked in detail: this Office is a volume licence without a
  Microsoft 365 sign-in, and the Trust Center dialog only accepts HTTPS catalog URLs, so no local
  deployment route exists.
- **The pane cannot reach anything at all** on some Office builds, and the loopback exemption is missing.
  That exemption was only needed for the old dev-server setup; if you hit it, enable loopback for the
  Office webview once in an **elevated** PowerShell and restart PowerPoint:

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
2. Install Script Lab inside PowerPoint (`Insert > Get Add-ins`, search for `Script Lab`).

### Start the local helper

The helper installed in step 1 already listens on `http://localhost:8000`, so there is nothing to start
while the add-in is installed. (Without the add-in, run `install.ps1 -NoAutostart`, or start the helper
in a console with `run-server.cmd`: `ZoteroHelper.exe --no-static` serves only the JSON API.)

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

## 3. Development notes

### Repository structure

- Root files `index.html`, `style.css`, and `script.js` are for Script Lab usage.
- `zotero-addon/www/` is the add-in web root: `taskpane.html`, `commands.html`, `style.css`, `frontend_core.js`.
- `zotero-addon/helper/ZoteroHelper.cs` is the whole helper (API proxy + HTTPS file server), compiled by `install.ps1`.
- `zotero-addon/manifest.xml` is the add-in manifest; `install.ps1` copies it to `%LOCALAPPDATA%\ZoteroCitations`.
- `shared/frontend_core.js` is the single source of truth for frontend behavior.
- `install.ps1` / `uninstall.ps1` install and remove the add-in on Windows.

### Generated and wrapped files

These files should not be treated as primary edit targets:

- `script.js`
- `zotero-addon/www/frontend_core.js`
- `zotero-addon/www/style.css`

For frontend behavior changes, edit `shared/frontend_core.js`; for styling, edit `style.css`; then
regenerate the outputs.

### Regenerating shared frontend files

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\sync_shared.ps1
```

This updates `script.js`, `zotero-addon/www/frontend_core.js`, and `zotero-addon/www/style.css`.
Afterwards re-run `install.ps1` so PowerPoint gets the new files.

### Validation

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\test-helper.ps1
node tools/test-frontend.js
```

The first compiles the helper, starts it against a stub Better BibTeX and checks the API, the CORS
preflight, the HTTPS file server, the no-cache headers and path traversal handling (29 checks). The
second runs the pane's code without a browser: the citation text builder (picker locator, prefix,
suffix, suppress author), the HTML-to-formatting-runs parser, and an Office.js usage lint that fails
when a collection is read before its `load()` was delivered by an awaited `context.sync()` - the bug
that made Generate Bibliography fail once in PowerPoint. The lint ships with fixtures that must be
flagged (they are part of the test), and reintroducing the load/sync bug in the frontend makes it fail
on the real file; the bibliography entry splitter and the character-budget fallback are covered too (39 checks). Then restart PowerPoint
and open the pane to check the add-in itself.

### Implementation notes

- `ZoteroHelper.exe` serves the JSON API on `http://localhost:8000` and the pane on `https://localhost:23000`.
- `/zotero` proxies Better BibTeX CAYW calls (the Zotero picker can stay open for minutes, so the
  upstream call has a 10 minute ceiling).
- `/bibliography` proxies Better BibTeX JSON-RPC bibliography generation; `format: "html"` in the
  request becomes `contentType: html` upstream, which is what makes real italics/bold/sub/superscript
  possible in the References slide.
- `/health` is used by the UI status indicator; `POST /log` lets the pane write its own errors into `server.log`.
- Written in C# 5 on purpose: the compiler that ships with Windows is not a Roslyn compiler.
- Bibliography output is based on citation keys stored in slide metadata.

### Useful links

- https://retorque.re/zotero-better-bibtex/citing/cayw
- https://retorque.re/zotero-better-bibtex/exporting/json-rpc/index.html
- https://www.zotero.org/support/dev/web_api/v3/basics
- https://learn.microsoft.com/en-us/javascript/api/powerpoint
