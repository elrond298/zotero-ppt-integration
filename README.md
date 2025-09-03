# PowerPoint Zotero Integration Add-in

This project provides a simple [Script Lab](https://aka.ms/scriptlab)-based add-in for Microsoft PowerPoint. It allows you to connect to your running Zotero instance to insert in-text citations and generate a complete bibliography on a new slide.

The add-in works by communicating with a local Python server, which in turn communicates with the [Better BibTeX for Zotero](https://retorque.re/zotero-better-bibtex/) (BBT) plugin.

## Features

*   **Add In-Text Citations**: Pull up the Zotero citation picker directly from PowerPoint to insert formatted citations like `(Author, Year)` onto your slides.
*   **Track Citations per Slide**: The add-in stores citation keys in the metadata of each slide, keeping your references organized.
*   **Manage Slide Citations**: View a list of all citations on the currently selected slide and remove any that are no longer needed.
*   **Generate Bibliography**: Automatically collect all unique citations from your entire presentation and generate a formatted bibliography on a new, dedicated "References" slide.

## 1. Pre-requisites

Before you begin, ensure you have the following software installed and running:

1.  **Zotero**: The Zotero desktop application must be installed and running in the background. You can download it from [zotero.org](https://www.zotero.org/).
2.  **Better BibTeX for Zotero (BBT)**: This essential Zotero plugin provides the local API that the add-in communicates with.
    *   Go to the [latest BBT release page](https://github.com/retorquere/zotero-better-bibtex/releases/latest).
    *   Download the `.xpi` file.
    *   In Zotero, go to `Tools > Add-ons`, click the gear icon ⚙️, and select "Install Add-on From File..." to install the downloaded `.xpi` file.
    *   Restart Zotero.

3.  **Python 3**: The backend server is a Python script. Make sure you have Python 3 installed on your system. You can download it from [python.org](https://www.python.org/downloads/).

## 2. Install the Script Lab Add-in

[Script Lab](https://aka.ms/scriptlab) is a fantastic tool for developing and running Office Add-ins directly within the Office application.

1.  Open **PowerPoint**.
2.  Go to the **Insert** tab and click **Get Add-ins**.
3.  Search for "Script Lab" and click **Add**.
4.  Once installed, you will see a new **Script Lab** tab in the PowerPoint ribbon.

## 3. Create the Add-in Snippet

Now, you will copy the project files into a new Script Lab snippet.

1.  Go to the **Script Lab** tab and click **Code**. This will open the Script Lab task pane.
2.  Click the menu icon (☰) in the top left of the task pane and select **New**.
3.  You will see several tabs: `Script`, `HTML`, `CSS`, and `Libraries`. You need to copy and paste the contents of the provided files into the corresponding tabs.

#### a. HTML Tab

Copy the entire content of `index.html` and paste it into the **HTML** tab, replacing any existing content.

```html
<button id="add-citation" class="ms-Button">
		<span class="ms-Button-label">Add Citation</span>
</button>
<button id="generate-bibliography" class="ms-Button">
    <span class="ms-Button-label">Generate Bibliography</span>
</button>
<!-- Changed from <pre> to <div> to hold interactive list -->
<div id="output"></div>
```

#### b. CSS Tab

Copy the entire content of `style.css` and paste it into the **CSS** tab.

```css
section.samples {
    margin-top: 20px;
}

section.samples .ms-Button, section.setup .ms-Button {
    display: block;
    margin-bottom: 5px;
    margin-left: 20px;
    min-width: 80px;
}

#output p {
    margin-bottom: 5px;
    font-weight: bold;
}

#output ul {
    list-style-type: none;
    padding-left: 0;
    margin-top: 5px;
}

#output li {
    display: flex;
    align-items: center;
    margin-bottom: 4px;
    font-family: Consolas, monaco, monospace;
    font-size: 14px;
}

.remove-btn {
    background-color: #f44336; /* Red */
    color: white;
    border: none;
    border-radius: 50%;
    cursor: pointer;
    width: 20px;
    height: 20px;
    font-size: 14px;
    line-height: 20px;
    text-align: center;
    margin-right: 8px;
    padding: 0;
    font-weight: bold;
    flex-shrink: 0; /* Prevents the button from shrinking */
}

.remove-btn:hover {
    background-color: #d32f2f; /* Darker red */
}
```

#### c. Script Tab

Copy the entire content of `script.js` and paste it into the **Script** tab.

```javascript
// --- Configuration ---
const PROXY_ENDPOINT = "http://localhost:8000/zotero";
const BIB_ENDPOINT = "http://localhost:8000/bibliography";
const ZOTERO_TAG_KEY = "ZOTERO_CITATION_KEYS"; // Key for storing citation data in slide's custom properties

// --- Office.onReady Initialization ---
// ... (rest of the script.js file) ...
```

#### d. Libraries Tab

This project does not require any external libraries, so you can leave the **Libraries** tab empty.

Finally, give your snippet a name by clicking on "Untitled Snippet" at the top of the pane.

## 4. Run the Python Server

The add-in cannot talk to Zotero directly due to browser security policies (CORS). This Python script acts as a simple local server to bridge the communication.

1.  Save the Python code into a file named `server.py` on your computer.
2.  Open a terminal or command prompt.
3.  Navigate to the directory where you saved `server.py`.
4.  Run the server with the following command:

    ```bash
    python server.py
    ```

5.  You should see a message like `Proxy server running on http://localhost:8000`.

**Important**: This server must remain running in the background whenever you are using the PowerPoint add-in.

## 5. Using the Add-in

With the server running and the snippet set up in Script Lab, you are ready to use the add-in.

1.  In the Script Lab task pane, click the **Run** button (▶).
2.  The add-in interface with the "Add Citation" and "Generate Bibliography" buttons will appear in the task pane.

#### Adding a Citation

1.  Click on a slide where you want to add a citation. If your cursor is inside a text box, the citation will be appended there. Otherwise, a new text box will be created.
2.  Click the **Add Citation** button in the add-in pane.
3.  The Zotero citation picker window will appear. Search for and select the reference(s) you want to cite.
4.  Press Enter. The formatted in-text citation will be added to your slide.

#### Generating the Bibliography

1.  After adding all your citations throughout the presentation, click the **Generate Bibliography** button.
2.  The add-in will scan every slide, collect all unique citation keys, and request a formatted bibliography from Zotero.
3.  A new slide titled "References" will be created at the end of your presentation containing the full bibliography.

## 6. Managing Citations

The add-in provides a simple way to see and manage the citations associated with each slide.

*   **Viewing Citations**: When you select a slide, the add-in pane will automatically update to show a list of all Zotero citation keys stored on that slide (e.g., `correia2021`, `doeEtAl2023`).
*   **Removing a Citation**: If you delete a citation from the slide text, its key will still be stored in the slide's metadata. To remove it completely (so it doesn't appear in the bibliography), click the red **×** button next to the citation key in the add-in pane. This will remove the key from the slide's metadata.
*  **Bibliography is generated based on metadata alone**, delete citations manually if necessary.

This ensures that your final bibliography is always accurate and reflects only the citations present in your presentation.

## 7. Useful links
https://retorque.re/zotero-better-bibtex/citing/cayw
https://retorque.re/zotero-better-bibtex/exporting/json-rpc/index.html
https://www.zotero.org/support/dev/web_api/v3/basics
https://learn.microsoft.com/en-us/javascript/api/powerpoint