// --- Configuration ---
const PROXY_ENDPOINT = "http://localhost:8000/zotero";
const BIB_ENDPOINT = "http://localhost:8000/bibliography";
const ZOTERO_TAG_KEY = "ZOTERO_CITATION_KEYS"; // Key for storing citation data in slide's custom properties

// --- Office.onReady Initialization ---
// This function is called when the Office host is ready.
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    // Wire up the button event listeners
    document.getElementById("add-citation").addEventListener("click", handleAddCitation);
    document.getElementById("generate-bibliography").addEventListener("click", handleGenerateBibliography);

    // Add a click listener to the output area for handling tag removal (event delegation)
    document.getElementById("output").addEventListener("click", handleRemoveClick);

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
async function handleAddCitation() {
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

/**
 * Handles clicks within the output area to delegate removal actions.
 * @param {Event} event The click event.
 */
function handleRemoveClick(event) {
  // Find the closest ancestor that is a remove button
  const removeButton = event.target.closest(".remove-btn");
  if (removeButton) {
    const keyToRemove = removeButton.dataset.key;
    if (keyToRemove) {
      removeCitation(keyToRemove);
    }
  }
}

/**
 * Gathers all unique citation keys from the presentation, sends them to the
 * server for formatting, and adds the resulting bibliography to a new slide.
 */
async function handleGenerateBibliography() {
  const outputElement = document.getElementById("output");
  outputElement.textContent = "Generating bibliography...";

  try {
    // Use a Set to automatically handle duplicates
    const allCitationKeys = new Set();

    // 1. Loop over all slides and collect citation keys
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      // This is more efficient than syncing inside the loop.
      // First, request to load the tags for all slides.
      for (const slide of slides.items) {
        slide.tags.load("key, value");
      }
      await context.sync();

      // Now that all tags are loaded, process them
      for (const slide of slides.items) {
        const zoteroTag = slide.tags.items.find(tag => tag.key === ZOTERO_TAG_KEY);
        if (zoteroTag) {
          try {
            const keysArray = JSON.parse(zoteroTag.value);
            if (Array.isArray(keysArray)) {
              keysArray.forEach(key => allCitationKeys.add(key));
            }
          } catch (e) {
            console.warn(`Could not parse citation keys on a slide`, e);
          }
        }
      }
    });

    const uniqueKeys = Array.from(allCitationKeys);
    console.log("Found unique citation keys:", uniqueKeys);

    if (uniqueKeys.length === 0) {
      outputElement.textContent = "No citations found in the presentation.";
      return;
    }

    // 2. Send keys to the server to get the formatted bibliography
    const formattedBibliography = await fetchBibliographyFromServer(uniqueKeys);

    // 3. Add the bibliography to a new slide
    await addBibliographySlide(formattedBibliography);

    outputElement.textContent = "Bibliography generated successfully!";

  } catch (error) {
    console.error("Error generating bibliography:", error);
    outputElement.textContent = "Error: Could not generate bibliography.";
  }
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

/**
 * Sends citation keys to the server and returns the formatted bibliography string.
 * @param {string[]} keys - An array of unique citation keys.
 * @returns {Promise<string>} A promise that resolves to the formatted bibliography text.
 */
async function fetchBibliographyFromServer(keys) {
  const response = await fetch(BIB_ENDPOINT, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    // You can configure the style here, e.g., 'apalike', 'unsrt', 'plain'
    body: JSON.stringify({ keys: keys, style: 'apalike' })
  });

  if (!response.ok) {
    const errorData = await response.json();
    throw new Error(errorData.error || `Server responded with status: ${response.status}`);
  }

  const data = await response.json();
  return data.bibliography;
}

// --- Core PowerPoint Interaction Functions ---

/**
 * Inserts citations into the presentation. If the cursor is in a text box,
 * it inserts the text there. Otherwise, it creates a new text box for the citation.
 * It also stores citation keys in the slide's custom tags and updates the UI.
 * @param {Array} zoteroItems An array of Zotero item objects from the API.
 */
async function insertCitationsIntoPowerPoint(zoteroItems) {
  try {
    await PowerPoint.run(async (context) => {
      // Get the currently selected slide. This is needed for both scenarios.
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

      // --- 3. Insert the Citation Text into the Slide ---
      try {
        // SCENARIO 1: Try to insert text at the current cursor position.
        // This will succeed if the cursor is in a text box.
        const selectedTextRange = context.presentation.getSelectedTextRange();
        selectedTextRange.load("text"); // We must load a property to validate the object.
        await context.sync();

        // Insert the text. 'Replace' is ideal: it inserts at the cursor if nothing
        // is selected, or replaces the currently highlighted text.
        const originalText = selectedTextRange.text;
        selectedTextRange.text = originalText + ' ' + citationsText;
        await context.sync();

      } catch (error) {
        // SCENARIO 2: The cursor is not in a text box.
        // Check if the error is the specific one we expect for no text selection.
        if (error.name === 'RichApi.Error' && error.code === 'GeneralException') {
          // Fallback: Create a new text box on the slide.
          console.log("No text range selected. Creating a new text box.");
          const leftPosition = 100;
          const topPosition = 150;
          const textBox = slide.shapes.addTextBox(citationsText, { left: leftPosition, top: topPosition, width: 400, height: 50 });
          textBox.load("textFrame/textRange"); // Good practice to load the new object.
          await context.sync();
        } else {
          // For any other unexpected error, re-throw it to be caught by the outer handler.
          console.error("An unexpected error occurred during text insertion:", error);
          throw error;
        }
      }
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
 * Removes a specific Zotero citation key from the currently selected slide's custom tags.
 * @param {string} keyToRemove The citation key to be removed.
 */
async function removeCitation(keyToRemove) {
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const customTags = slide.tags;
      customTags.load("key, value");
      await context.sync();

      const zoteroTag = customTags.items.find(tag => tag.key === ZOTERO_TAG_KEY);

      if (zoteroTag) {
        let currentKeys = [];
        try {
          currentKeys = JSON.parse(zoteroTag.value);
        } catch (e) {
          console.error("Could not parse existing tags for removal:", e);
          return; // Exit if tags are corrupt
        }

        // Filter out the key to be removed
        const updatedKeys = currentKeys.filter(key => key !== keyToRemove);

        // Update the tag with the new array
        slide.tags.add(ZOTERO_TAG_KEY, JSON.stringify(updatedKeys));
        await context.sync();
      }
    });

    // Refresh the UI to show the change
    await displayCitationsFromSlide();

  } catch (error) {
    console.error("Error removing citation:", error);
    document.getElementById("output").textContent = "Error: Could not remove citation.";
  }
}

/**
 * Reads the Zotero citation keys from the currently selected slide's custom tags
 * and displays them as an interactive list in the add-in's UI.
 */
async function displayCitationsFromSlide() {
  const outputElement = document.getElementById("output");
  // Clear previous content
  outputElement.innerHTML = "";

  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      if (!slide) {
        outputElement.textContent = "Select a slide to view its citations.";
        return;
      }

      const customTags = slide.tags;
      customTags.load("key, value");
      await context.sync();

      const zoteroTag = customTags.items.find(tag => tag.key === ZOTERO_TAG_KEY);

      if (zoteroTag) {
        try {
          const keysArray = JSON.parse(zoteroTag.value);
          if (Array.isArray(keysArray) && keysArray.length > 0) {
            // Create a header and a list
            const header = document.createElement("p");
            header.textContent = "Citations on this slide:";
            const list = document.createElement("ul");

            keysArray.forEach(key => {
              const listItem = document.createElement("li");

              // Create the remove button
              const removeButton = document.createElement("button");
              removeButton.className = "remove-btn";
              removeButton.innerHTML = "&times;"; // 'x' symbol
              removeButton.title = `Remove citation: ${key}`;
              removeButton.dataset.key = key; // Store the key in a data attribute

              // Create the text span
              const keyText = document.createElement("span");
              keyText.textContent = key;

              // Assemble the list item
              listItem.appendChild(removeButton);
              listItem.appendChild(keyText);
              list.appendChild(listItem);
            });

            outputElement.appendChild(header);
            outputElement.appendChild(list);

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
    if (outputElement.innerHTML === "") {
      outputElement.textContent = "Could not load citations for this slide.";
    }
  }
}

/**
 * Adds a new bibliography slide to the end of the presentation.
 * @param bibliographyText The string content for the bibliography.
 */
async function addBibliographySlide(bibliographyText) {
  await PowerPoint.run(async function (context) {
    context.presentation.slides.add();
    await context.sync();

    // The index for the new slide will be the current number of slides.
    const newSlideIndex = context.presentation.slides.getCount();
    await context.sync();

    // Get the newly added slide by its index.
    context.presentation.load("slides");
    await context.sync();
    const newSlide = context.presentation.slides.getItemAt(newSlideIndex.value-1);
    newSlide.load("id");
    await context.sync();
    // No need to load 'id' here unless you specifically need it later.
    // Operations on the slide object will work without loading a property first.

    // Add a title
    const titleShape = newSlide.shapes.addTextBox("References", {
      left: 50,
      top: 50,
      width: 860,
      height: 100
    });
    titleShape.textFrame.textRange.font.size = 44;

    // Add the bibliography content
    const contentShape = newSlide.shapes.addTextBox(bibliographyText, {
      left: 50,
      top: 150,
      width: 860,
      height: 350
    });
    contentShape.textFrame.textRange.font.size = 14;
    // Allow the text box to resize vertically to fit all the content
    contentShape.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeShapeToFitText;

    // Sync all the queued changes (add slide, add text boxes, format text).
    await context.sync();
  });
}

