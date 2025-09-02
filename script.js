// --- Configuration ---
const PROXY_ENDPOINT = "http://localhost:8000/zotero";
const ZOTERO_TAG_KEY = "ZOTERO_CITATION_KEYS"; // Key for storing citation data in slide's custom properties

// --- Office.onReady Initialization ---
// This function is called when the Office host is ready.
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    // Wire up the 'Run' button event listener
    document.getElementById("run").addEventListener("click", run);

    // Add an event handler for when the slide selection changes.
    // This will keep the displayed citation list in sync with the selected slide.
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      displayCitationsFromSlide,
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Could not register selection change handler: " + asyncResult.error.message);
        }
      }
    );

    // Perform an initial load of citations for the currently selected slide.
    displayCitationsFromSlide();
  }
});


// --- Event Handlers ---

/**
 * Fetches citation data from the Zotero proxy and triggers the insertion process.
 */
async function run() {
  var xhr = new XMLHttpRequest();
  xhr.open("GET", PROXY_ENDPOINT, true);
  xhr.onreadystatechange = function () {
    if (xhr.readyState === 4) {
      if (xhr.status === 200) {
        try {
          var response = JSON.parse(xhr.responseText);
          console.log("Zotero response:", response);

          // Call the function to insert citations into PowerPoint
          insertCitationsIntoPowerPoint(response);

        } catch (e) {
          console.error("Error parsing JSON response:", e);
          document.getElementById("output").textContent = "Error: Could not parse Zotero data.";
        }
      } else {
        console.error("Request failed with status:", xhr.status);
        document.getElementById("output").textContent = `Error: Could not connect to Zotero (Status: ${xhr.status}).`;
      }
    }
  };
  xhr.onerror = function () {
    console.error("Request failed");
    document.getElementById("output").textContent = "Error: Request to Zotero proxy failed. Is it running?";
  };
  xhr.send();
}

// --- Helper Functions ---

/**
 * Formats a single Zotero item into an "Author, Year" string.
 * Handles 1, 2, or 3+ authors.
 * @param {object} item A single Zotero item object from the API response.
 * @returns {string} A formatted citation string, e.g., "Correia, 2021" or "Doe et al., 2023".
 */
function formatSingleCitation(item) {
  const creators = item.item.creators;
  const date = item.item.date || "";
  const year = date.substring(0, 4) || "n.d."; // n.d. for "no date"

  let authorString = "Unknown Author";
  if (creators && creators.length > 0) {
    if (creators.length === 1) {
      authorString = creators[0].lastName;
    } else if (creators.length === 2) {
      authorString = `${creators[0].lastName} & ${creators[1].lastName}`;
    } else {
      authorString = `${creators[0].lastName} et al.`;
    }
  }
  return `${authorString}, ${year}`;
}


// --- Core PowerPoint Interaction Functions ---

/**
 * Inserts citations as a hyperlinked text box, stores citation keys in the slide's custom tags,
 * and then updates the UI to show all keys for the slide.
 * @param {Array} zoteroItems An array of Zotero item objects from the API.
 */
async function insertCitationsIntoPowerPoint(zoteroItems) {
  try {
    await PowerPoint.run(async (context) => {
      // Get the currently selected slide.
      const slides = context.presentation.getSelectedSlides();
      slides.load("items");
      await context.sync();

      if (slides.items.length === 0) {
        console.error("No slide selected.");
        document.getElementById("output").textContent = "Please select a slide first.";
        return;
      }
      const slide = slides.items[0];

      // Check if the response is a valid array
      if (!Array.isArray(zoteroItems) || zoteroItems.length === 0) {
        console.log("No valid citation items found in the response.");
        document.getElementById("output").textContent = "No citations found to insert.";
        return;
      }

      // --- 1. Update Slide Custom Properties (Tags) ---
      const customTags = slide.tags;
      customTags.load("key, value");
      await context.sync();

      const existingTag = customTags.items.find(tag => tag.key === ZOTERO_TAG_KEY);
      let citationKeys = new Set();

      if (existingTag) {
        try {
          const keysArray = JSON.parse(existingTag.value);
          if (Array.isArray(keysArray)) {
            keysArray.forEach(key => citationKeys.add(key));
          }
        } catch (e) {
          console.error("Could not parse existing citation tags:", e);
        }
      }
      // Add new keys from the current insertion
      zoteroItems.forEach(item => citationKeys.add(item.citationKey));

      // Add or update the tag on the slide
      const updatedKeysArray = Array.from(citationKeys);
      slide.tags.add(ZOTERO_TAG_KEY, JSON.stringify(updatedKeysArray));

      // --- 2. Create and Insert the Citation Text Box ---
      const citationParts = zoteroItems.map(item => formatSingleCitation(item));
      const citationsText = `(${citationParts.join('; ')})`;

      const leftPosition = 100;
      const topPosition = 150;
      const textBox = slide.shapes.addTextBox(citationsText, { left: leftPosition, top: topPosition, width: 400, height: 50 });
      textBox.load("textFrame/textRange");
      await context.sync();
    });

    // --- 4. Update the UI to reflect the changes ---
    // This will read the tags we just wrote and display them.
    await displayCitationsFromSlide();

  } catch (error) {
    console.error("Error interacting with PowerPoint:", error);
    document.getElementById("output").textContent = "Error: Could not insert citations.";
  }
}

/**
 * Reads the Zotero citation keys from the currently selected slide's custom tags
 * and displays them in the add-in's UI.
 */
async function displayCitationsFromSlide() {
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      if (!slide) {
        // This can happen if no slides exist or none are selected.
        document.getElementById("output").textContent = "Select a slide to view its citations.";
        return;
      }

      const customTags = slide.tags;
      customTags.load("key, value");
      await context.sync();

      const zoteroTag = customTags.items.find(tag => tag.key === ZOTERO_TAG_KEY);
      const outputElement = document.getElementById("output");

      if (zoteroTag) {
        try {
          const keysArray = JSON.parse(zoteroTag.value);
          if (Array.isArray(keysArray) && keysArray.length > 0) {
            outputElement.textContent = `Citations on this slide:\n${keysArray.join('\n')}`;
          } else {
            outputElement.textContent = "No Zotero citations found on this slide.";
          }
        } catch (e) {
          console.error("Error parsing citation tags from slide:", e);
          outputElement.textContent = "Error reading citations from this slide.";
        }
      } else {
        outputElement.textContent = "No Zotero citations found on this slide.";
      }
    });
  } catch (error) {
    console.error("Error displaying citations from slide:", error);
    // Avoid overwriting an important message if this fails in the background.
  }
}
