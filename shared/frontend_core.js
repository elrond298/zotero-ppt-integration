const PROXY_ENDPOINT = "http://localhost:8000/zotero";
const BIB_ENDPOINT = "http://localhost:8000/bibliography";
const HEALTH_ENDPOINT = "http://localhost:8000/health";
const LOG_ENDPOINT = "http://localhost:8000/log";
const SNAPSHOT_ENDPOINT = "http://localhost:8000/snapshot";
const DEFAULT_BIBLIOGRAPHY_STYLE = "journal-of-geophysical-research-atmospheres";
const BIBLIOGRAPHY_STYLE_STORAGE_KEY = "zotero-ppt:bibliography-style";
const ZOTERO_TAG_KEY = "ZOTERO_CITATION_KEYS";
/* Invisible separator/plus characters written by an earlier version of the add-in: PowerPoint shows
   them, so they are no longer written, but decks already carrying them are cleaned up. */
const CITATION_MARK_START = "\u2063";
const CITATION_MARK_END = "\u2064";
const ZOTERO_BIBLIOGRAPHY_TAG = "ZOTERO_BIBLIOGRAPHY";
const BIBLIOGRAPHY_TITLE = "References";
const BIBLIOGRAPHY_CONTINUATION_TITLE = "References (cont.)";
const BIBLIOGRAPHY_SPLIT_STORAGE_KEY = "zotero-ppt:split-bibliography";
const DEFAULT_SPLIT_BIBLIOGRAPHY = true;
const BIBLIOGRAPHY_FONT_SIZE_STORAGE_KEY = "zotero-ppt:bibliography-font-size";
const BIBLIOGRAPHY_FONT_SIZES = [10, 11, 12, 14, 16, 18, 20, 24];
const DEFAULT_BIBLIOGRAPHY_FONT_SIZE = 14;
/* Fallback when PowerPoint does not report a text height that reacts to the text: about 120
   characters fit per line and about 15 lines fit on a slide, so this leaves headroom. */
const BIBLIOGRAPHY_CHARS_PER_SLIDE = 1800;
/* Set to false when the host cannot move slides or render them as images
   (Slide.moveTo and Slide.getImageAsBase64 arrived with PowerPointApi 1.8). */
let canMoveSlides = supportsPowerPointApi("1.8");
let canSnapshotSlides = null;

const LOG_LEVELS = {
  NONE: 0,
  ERROR: 1,
  WARN: 2,
  INFO: 3,
  DEBUG: 4,
};

const LOG_LEVEL = LOG_LEVELS.WARN;

function logDebug(...args) {
  if (LOG_LEVEL >= LOG_LEVELS.DEBUG) {
    console.log("[DEBUG]", ...args);
  }
}

function logInfo(...args) {
  if (LOG_LEVEL >= LOG_LEVELS.INFO) {
    console.log("[INFO]", ...args);
  }
}

function logWarn(...args) {
  if (LOG_LEVEL >= LOG_LEVELS.WARN) {
    console.warn("[WARN]", ...args);
  }
}

function logError(...args) {
  if (LOG_LEVEL >= LOG_LEVELS.ERROR) {
    console.error("[ERROR]", ...args);
  }
}

function log(...args) {
  logInfo(...args);
}

let healthIntervalId = null;

/* A pane that loaded while the helper was restarting can end up without its stylesheet; in that
   case the page reloads itself once. */
(function reloadIfStylesheetFailed() {
  const stylesheet = document.querySelector('link[rel="stylesheet"]');
  const alreadyReloaded = window.sessionStorage.getItem("zotero-ppt:stylesheet-reload") === "1";
  if (stylesheet && !stylesheet.sheet && !alreadyReloaded) {
    window.sessionStorage.setItem("zotero-ppt:stylesheet-reload", "1");
    window.location.reload();
  }
})();

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) {
    logWarn("This add-in is built for PowerPoint. Current host:", info.host);
    return;
  }

  const wireEvents = () => {
    const addBtn = document.getElementById("add-citation");
    const addSelectedBtn = document.getElementById("add-citation-selected");
    const bibBtn = document.getElementById("generate-bibliography");
    const bibliographyStyleSelect = document.getElementById("bibliography-style");
    const output = document.getElementById("output");

    if (!addBtn || !bibBtn || !output) {
      logError("UI elements not found in the DOM. Check that scripts load after HTML body.");
      return;
    }

    addBtn.addEventListener("click", handleAddCitation);
    if (addSelectedBtn) {
      addSelectedBtn.addEventListener("click", handleAddCitationSelected);
    }
    bibBtn.addEventListener("click", handleGenerateBibliography);
    if (bibliographyStyleSelect) {
      initializeBibliographyStyleSelect(bibliographyStyleSelect);
      bibliographyStyleSelect.addEventListener("change", handleBibliographyStyleChange);
    }
    const fontSizeSelect = document.getElementById("bibliography-font-size");
    if (fontSizeSelect instanceof HTMLSelectElement) {
      initializeBibliographyFontSize(fontSizeSelect);
      fontSizeSelect.addEventListener("change", handleBibliographyFontSizeChange);
    }
    const allCitationsElement = document.getElementById("all-citations");
    if (allCitationsElement) {
      allCitationsElement.addEventListener("click", handleJumpClick);
      // PowerPoint's slide collection can still be empty while the pane is wiring itself up, so the
      // first scan of the whole deck waits a moment; edits refresh it right away afterwards.
      window.setTimeout(displayAllCitations, 1500);
    }
    const splitBibliographyInput = document.getElementById("split-bibliography");
    if (splitBibliographyInput instanceof HTMLInputElement) {
      initializeBibliographySplitToggle(splitBibliographyInput);
      splitBibliographyInput.addEventListener("change", handleBibliographySplitChange);
    }
    output.addEventListener("click", handleRemoveClick);

    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      displayCitationsFromSlide,
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          logError("Could not register selection change handler: " + asyncResult.error.message);
        }
      }
    );

    displayCitationsFromSlide();
    startHealthPolling();
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", wireEvents, { once: true });
  } else {
    wireEvents();
  }
});

async function handleAddCitation() {
  const xhr = new XMLHttpRequest();
  xhr.open("GET", PROXY_ENDPOINT, true);
  xhr.onreadystatechange = function () {
    if (xhr.readyState === 4) {
      if (xhr.status === 200) {
        try {
          const response = JSON.parse(xhr.responseText);
          log("Zotero response:", response);
          insertCitationsIntoPowerPoint(response);
        } catch (error) {
          logError("Error parsing JSON response:", error);
          document.getElementById("output").textContent = "Error: Could not parse Zotero data.";
        }
      } else {
        logError("Request failed with status:", xhr.status);
        document.getElementById("output").textContent = `Error: Could not connect to Zotero (Status: ${xhr.status}).`;
      }
    }
  };
  xhr.onerror = function () {
    logError("Request failed");
    document.getElementById("output").textContent = "Error: Request to Zotero proxy failed. Is it running?";
  };
  xhr.send();
}

async function handleAddCitationSelected() {
  const xhr = new XMLHttpRequest();
  xhr.open("GET", PROXY_ENDPOINT + "?selected=true", true);
  xhr.onreadystatechange = function () {
    if (xhr.readyState === 4) {
      if (xhr.status === 200) {
        try {
          const response = JSON.parse(xhr.responseText);
          log("Zotero response (selected):", response);
          insertCitationsIntoPowerPoint(response);
        } catch (error) {
          logError("Error parsing JSON response:", error);
          document.getElementById("output").textContent = "Error: Could not parse Zotero data.";
        }
      } else {
        logError("Request failed with status:", xhr.status);
        document.getElementById("output").textContent = `Error: Could not connect to Zotero (Status: ${xhr.status}).`;
      }
    }
  };
  xhr.onerror = function () {
    logError("Request failed");
    document.getElementById("output").textContent = "Error: Request to Zotero proxy failed. Is it running?";
  };
  xhr.send();
}

async function checkHealthOnce() {
  const statusEl = document.getElementById("status");
  if (!statusEl) {
    return;
  }

  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 2500);
    const resp = await fetch(HEALTH_ENDPOINT, { signal: controller.signal });
    clearTimeout(timeout);

    if (resp.ok) {
      statusEl.textContent = "Zotero proxy connected";
      statusEl.className = "status ok";
    } else {
      statusEl.textContent = "Proxy unreachable (" + resp.status + ")";
      statusEl.className = "status error";
    }
  } catch (error) {
    statusEl.textContent = "Proxy not running";
    statusEl.className = "status error";
  }
}

function startHealthPolling() {
  const statusEl = document.getElementById("status");
  if (!statusEl) {
    return;
  }

  checkHealthOnce();
  if (healthIntervalId !== null) {
    return;
  }
  healthIntervalId = window.setInterval(checkHealthOnce, 8000);
}

function initializeBibliographyStyleSelect(selectElement) {
  const savedStyle = window.localStorage.getItem(BIBLIOGRAPHY_STYLE_STORAGE_KEY);
  const initialStyle = savedStyle || DEFAULT_BIBLIOGRAPHY_STYLE;
  let hasOption = false;
  for (let index = 0; index < selectElement.options.length; index += 1) {
    if (selectElement.options[index].value === initialStyle) {
      hasOption = true;
      break;
    }
  }
  selectElement.value = hasOption ? initialStyle : DEFAULT_BIBLIOGRAPHY_STYLE;
}

function initializeBibliographySplitToggle(inputElement) {
  const savedValue = window.localStorage.getItem(BIBLIOGRAPHY_SPLIT_STORAGE_KEY);
  inputElement.checked = savedValue === null ? DEFAULT_SPLIT_BIBLIOGRAPHY : savedValue === "true";
}

function handleBibliographySplitChange(event) {
  window.localStorage.setItem(BIBLIOGRAPHY_SPLIT_STORAGE_KEY, String(Boolean(event.target.checked)));
}

function isSplitBibliographyEnabled() {
  const inputElement = document.getElementById("split-bibliography");
  if (inputElement instanceof HTMLInputElement) return inputElement.checked;
  return DEFAULT_SPLIT_BIBLIOGRAPHY;
}

function handleBibliographyStyleChange(event) {
  const nextStyle = event.target.value || DEFAULT_BIBLIOGRAPHY_STYLE;
  window.localStorage.setItem(BIBLIOGRAPHY_STYLE_STORAGE_KEY, nextStyle);
}

function getSelectedBibliographyStyle() {
  const selectElement = document.getElementById("bibliography-style");
  if (selectElement instanceof HTMLSelectElement && selectElement.value) {
    return selectElement.value;
  }

  return window.localStorage.getItem(BIBLIOGRAPHY_STYLE_STORAGE_KEY) || DEFAULT_BIBLIOGRAPHY_STYLE;
}

function handleRemoveClick(event) {
  const removeButton = event.target.closest(".remove-btn");
  if (removeButton) {
    const keyToRemove = removeButton.dataset.key;
    if (keyToRemove) {
      removeCitation(keyToRemove);
    }
  }
}

/* Is this PowerPointApi version available in the host? Feature-detect instead of trying and
   failing: a rejected batch discards everything queued with it. */
function supportsPowerPointApi(version) {
  try {
    return Boolean(Office.context && Office.context.requirements && Office.context.requirements.isSetSupported("PowerPointApi", version));
  } catch (error) {
    return false;
  }
}

function capabilitySummary() {
  const versions = ["1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9"];
  const supported = versions.filter((version) => supportsPowerPointApi(version));
  const diagnostics = (Office.context && Office.context.diagnostics) || {};
  return "PowerPointApi " + (supported.length > 0 ? supported.join(",") : "none") +
    " | host=" + (diagnostics.host || "?") + " platform=" + (diagnostics.platform || "?") +
    " version=" + (diagnostics.version || "?");
}

/* Citations are stored as {"k": key, "l": label}; slides written before that hold plain key strings. */
function parseCitationTag(value) {
  const entries = [];
  let parsed;
  try {
    parsed = JSON.parse(value);
  } catch (error) {
    logWarn("Could not parse a citation tag", error);
    return entries;
  }
  if (!Array.isArray(parsed)) return entries;

  parsed.forEach((item) => {
    if (typeof item === "string" && item.length > 0) {
      entries.push({ key: item, label: "", shapeId: "", recorded: false });
    } else if (item && typeof item === "object" && typeof item.k === "string" && item.k.length > 0) {
      entries.push({
        key: item.k,
        label: typeof item.l === "string" ? item.l : "",
        shapeId: typeof item.s === "string" ? item.s : "",
        recorded: typeof item.s === "string" && item.s !== "",
      });
    }
  });
  return entries;
}

/* The label stored with a citation is the text that was written for it, but that text can be edited
   afterwards (a prefix added, "et al." inserted), so a citation counts as still present while its
   author and year are still somewhere on the slide. An entry without a label is always kept. */
function removeCitationText(slideText, label) {
  const text = String(slideText || "");
  const wanted = String(label || "").trim();
  if (wanted === "") return { text: text, removed: false };

  const escaped = wanted.replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/\s+/g, "\\s+");
  const match = new RegExp(escaped, "i").exec(text);
  if (!match) return { text: text, removed: false };

  let start = match.index;
  let end = match.index + match[0].length;
  const open = text.lastIndexOf("(", start);
  if (open >= 0 && text.slice(open + 1, start).trim() === "" && text.charAt(end) === ")") {
    start = open;
    end += 1;
  }

  const cleaned = tidyCitationText(text.slice(0, start) + text.slice(end));
  return { text: cleaned, removed: true };
}

/* Marker characters only. */
function stripMarkerJunk(text) {
  return String(text || "").replace(/[\u2063\u2064]/g, "");
}

/* Removes markers and the key text an earlier version of the add-in wrote into the citation, which
   PowerPoint showed as part of the citation. */
function repairCitationText(text, keys) {
  let result = String(text || "");
  (keys || []).forEach((key) => {
    if (!key) return;
    result = result.split(CITATION_MARK_START + key + CITATION_MARK_START).join("");
  });
  return stripMarkerJunk(result).replace(/\s{2,}/g, " ");
}

/* Finds the citation's text in a shape, tolerating the edits a user makes to it: the exact text
   when it is still there, otherwise the parenthesised group that still looks like it (same year, and
   as many of the same words as possible). A citation that was deleted leaves no such group. */
function findCitationSpan(text, label) {
  const haystack = String(text || "");
  const wanted = String(label || "").trim();
  if (wanted === "") return null;

  const exact = haystack.indexOf(wanted);
  if (exact >= 0) {
    // Take the parentheses around it with it ("(Smith, 2020, p. 42)"), but only when they hold
    // little else - "(see Smith, 2020)" is still the citation, "(Smith, 2020; Doe, 2019)" is not.
    const open = haystack.lastIndexOf("(", exact);
    const close = haystack.indexOf(")", exact + wanted.length);
    const group = open >= 0 && close >= 0 ? haystack.slice(open, close + 1) : "";
    const years = group.match(/\b(1[5-9]\d{2}|20\d{2})\b/g) || [];
    // The group may carry a little more than the label ("..., see page 2"), but not another
    // citation: those bring their own year and are left for their own removal.
    if (group !== "" && years.length <= 1 && close - open <= wanted.length + 40 && haystack.slice(open + 1, exact).trim().length <= 6) {
      return { start: open, end: close + 1 };
    }
    return { start: exact, end: exact + wanted.length };
  }

  const year = (wanted.match(/\b(1[5-9]\d{2}|20\d{2})\b/) || [])[0] || "";
  const words = (wanted.match(/[\p{L}][\p{L}'’.-]*/gu) || [])
    .filter((word) => word.length > 2 && word !== year)
    .map((word) => word.toLowerCase());

  const candidates = [];
  const pattern = /\([^()]*\)/g;
  let match = pattern.exec(haystack);
  while (match) {
    candidates.push({ start: match.index, end: match.index + match[0].length, text: match[0] });
    match = pattern.exec(haystack);
  }
  if (candidates.length === 0 && haystack.trim().length > 0 && haystack.length <= 200) {
    candidates.push({ start: 0, end: haystack.length, text: haystack });
  }

  let best = null;
  let bestScore = 0;
  candidates.forEach((candidate) => {
    const lower = candidate.text.toLowerCase();
    let score = year && lower.indexOf(year) >= 0 ? 2 : 0;
    score += words.filter((word) => lower.indexOf(word) >= 0).length;
    if (candidate.text.trim().length <= wanted.length + 8) score += 1;
    if (score > bestScore) {
      bestScore = score;
      best = candidate;
    }
  });
  return bestScore >= 2 ? best : null;
}

/* The id of the shape whose text holds what was just written. That id, stored on the slide's
   citation tag, is how the add-in knows later whether the citation's text box is still there. */
async function findCitationShapeId(context, slide, citationsText) {
  const shapes = slide.shapes;
  shapes.load("items/id,items/textFrame/textRange/text");
  await context.sync();

  const holders = shapes.items.filter((shape) => {
    const textFrame = shape.textFrame;
    const textRange = textFrame && textFrame.textRange ? textFrame.textRange : null;
    return (textRange && textRange.text ? textRange.text : "").indexOf(citationsText) >= 0;
  });
  const holder = holders[holders.length - 1];
  return holder ? String(holder.id || "") : "";
}

/* The shapes of this slide by id, with whether each still holds text. A citation recorded on a shape
   is still there while that shape exists with text - editing the citation's wording does not matter -
   and it is gone once the shape is gone or emptied. */
async function readCitationShapes(context, slide) {
  const byId = new Map();
  const shapes = slide.shapes;
  shapes.load("items/id,items/textFrame/textRange/text");
  await context.sync();
  shapes.items.forEach((shape) => {
    const textFrame = shape.textFrame;
    const textRange = textFrame && textFrame.textRange ? textFrame.textRange : null;
    const text = textRange && textRange.text ? textRange.text : "";
    byId.set(String(shape.id || ""), text.trim() !== "");
  });
  return byId;
}

/* Tidies the text a citation left behind: no double spaces, no separator without a citation. */
function tidyCitationText(text) {
  const cleaned = String(text || "")
    .replace(/\s{2,}/g, " ")
    .replace(/\(\s*[;,]\s*/g, "(")
    .replace(/\s*[;,]\s*\)/g, ")")
    .replace(/\(\s*\)/g, "")
  return cleaned.trim() === "" ? "" : cleaned.trim();
}

/* One entry per cited item, with the slides that cite it, for the pane's list. */
function groupCitationsByKey(citations) {
  const groups = new Map();
  citations.forEach((citation) => {
    const group = groups.get(citation.key) || { key: citation.key, label: citation.label, slides: [] };
    if (!group.label && citation.label) group.label = citation.label;
    if (group.slides.indexOf(citation.slideNumber) < 0) group.slides.push(citation.slideNumber);
    groups.set(citation.key, group);
  });
  return Array.from(groups.values());
}

/* {key, label, slideNumber} for every citation in the deck, in slide order. */
async function collectCitationsFromSlides() {
  const citations = [];

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    for (const slide of slides.items) {
      slide.tags.load("key, value");
    }
    await context.sync();

    let taggedSlides = 0;
    let sampleTag = "";
    slides.items.forEach((slide, index) => {
      const zoteroTag = slide.tags.items.find((tag) => tag.key === ZOTERO_TAG_KEY);
      if (!zoteroTag) return;
      taggedSlides += 1;
      if (!sampleTag) sampleTag = String(zoteroTag.value).slice(0, 120);
      parseCitationTag(zoteroTag.value).forEach((entry) => {
        citations.push({ key: entry.key, label: entry.label, slideNumber: index + 1 });
      });
    });
    reportToHelper(
      "citation scan: slides=" + slides.items.length + " tagged=" + taggedSlides +
        " citations=" + citations.length + " sample=" + sampleTag,
    );
  });

  return citations;
}

function initializeBibliographyFontSize(selectElement) {
  BIBLIOGRAPHY_FONT_SIZES.forEach((size) => {
    const option = document.createElement("option");
    option.value = String(size);
    option.textContent = size + " pt";
    selectElement.appendChild(option);
  });
  const saved = window.localStorage.getItem(BIBLIOGRAPHY_FONT_SIZE_STORAGE_KEY);
  const preferred = Number(saved);
  selectElement.value = String(
    BIBLIOGRAPHY_FONT_SIZES.indexOf(preferred) >= 0 ? preferred : DEFAULT_BIBLIOGRAPHY_FONT_SIZE,
  );
}

function handleBibliographyFontSizeChange(event) {
  const value = Number(event.target.value);
  const size = BIBLIOGRAPHY_FONT_SIZES.indexOf(value) >= 0 ? value : DEFAULT_BIBLIOGRAPHY_FONT_SIZE;
  window.localStorage.setItem(BIBLIOGRAPHY_FONT_SIZE_STORAGE_KEY, String(size));
}

function getBibliographyFontSize() {
  const selectElement = document.getElementById("bibliography-font-size");
  if (selectElement instanceof HTMLSelectElement) {
    const value = Number(selectElement.value);
    if (BIBLIOGRAPHY_FONT_SIZES.indexOf(value) >= 0) return value;
  }
  return DEFAULT_BIBLIOGRAPHY_FONT_SIZE;
}

/* Saves a PNG of a slide (PowerPointApi 1.8) through the helper, so the result of a generation can
   be looked at afterwards even when it failed. Unsupported hosts are asked only once. */
async function snapshotSlide(context, slide, name) {
  if (canSnapshotSlides === null) {
    canSnapshotSlides = supportsPowerPointApi("1.8");
    reportToHelper("capabilities: " + capabilitySummary());
  }
  if (!canSnapshotSlides) return;

  try {
    const image = slide.getImageAsBase64();
    await context.sync();
    if (image.value) {
      await fetch(SNAPSHOT_ENDPOINT + "?name=" + encodeURIComponent(name), { method: "POST", body: image.value });
    }
  } catch (error) {
    canSnapshotSlides = false;
    logWarn("Could not render the slide as an image", error);
    reportToHelper("slide.getImageAsBase64 failed" + describeError(error));
  }
}

async function snapshotCurrentSlide() {
  if (canSnapshotSlides === false) return;
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.getSelectedSlides();
      slides.load("items");
      await context.sync();
      if (slides.items.length > 0) await snapshotSlide(context, slides.items[0], "failure");
    });
  } catch (error) {
    logWarn("Could not snapshot the failing slide", error);
  }
}

/* Jumping needs Presentation.setSelectedSlides() with slide ids (PowerPointApi 1.5). */
async function jumpToSlide(slideNumber) {
  if (!supportsPowerPointApi("1.5")) return;
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();
      const slide = slides.items[slideNumber - 1];
      if (slide) context.presentation.setSelectedSlides([slide.id]);
      await context.sync();
    });
  } catch (error) {
    logWarn("Could not select slide", slideNumber, error);
  }
}

function handleJumpClick(event) {
  const row = event.target.closest("[data-slide-number]");
  if (!row) return;
  jumpToSlide(Number(row.dataset.slideNumber));
}

/* Every cited item with the slides that cite it; clicking a row jumps to the first one. */
async function displayAllCitations() {
  const container = document.getElementById("all-citations");
  if (!container) return;

  try {
    const groups = groupCitationsByKey(await collectCitationsFromSlides());
    if (groups.length === 0) {
      container.textContent = "No citations in this presentation yet.";
      return;
    }

    const list = document.createElement("ul");
    groups.forEach((group) => {
      const listItem = document.createElement("li");
      listItem.className = "citation-row";
      listItem.dataset.slideNumber = String(group.slides[0]);
      listItem.title = "Go to slide " + group.slides[0];

      const label = document.createElement("span");
      label.className = "citation-label";
      label.textContent = group.label || group.key;

      const slides = document.createElement("span");
      slides.className = "citation-slides";
      slides.textContent = group.slides.join(", ");

      listItem.appendChild(label);
      listItem.appendChild(slides);
      list.appendChild(listItem);
    });
    container.replaceChildren(list);
  } catch (error) {
    logWarn("Could not list the citations of this presentation", error);
    reportToHelper("listing all citations failed" + describeError(error));
    container.textContent = "Could not read the citations of this presentation.";
  }
}

/* A key that BBT cannot resolve (renamed or deleted item) renders as nothing, so every key is
   checked on its own. A failed check is not reported as a missing key. */
async function findMissingCitationKeys(keys, style) {
  const results = await Promise.all(
    keys.map(async (key) => {
      try {
        const text = await fetchBibliographyFromServer([key], style, "text");
        return typeof text === "string" && text.trim().length > 0 ? null : key;
      } catch (error) {
        logWarn("Could not check citation key", key, error);
        return null;
      }
    }),
  );
  return results.filter((key) => key !== null);
}

/* Runs one step of the bibliography flow, so a failure can name the step it happened in. */
async function runStep(label, action) {
  try {
    return await action();
  } catch (error) {
    if (error && !error.step) error.step = label;
    throw error;
  }
}

/* The pane reports its own failures to the helper, so they show up in server.log. */
function reportToHelper(message) {
  try {
    fetch(LOG_ENDPOINT, { method: "POST", body: String(message), keepalive: true }).catch(() => {});
  } catch (error) {
    // Diagnostics must never break the add-in.
  }
}

/* Office.js errors carry the failing statement in debugInfo, which is what makes them fixable. */
function describeError(error) {
  if (!error) return "";
  const parts = [];
  if (error.step) parts.push(error.step);
  if (error.code) parts.push(String(error.code));
  const debugInfo = error.debugInfo;
  if (debugInfo) {
    if (debugInfo.statement) {
      parts.push("`" + String(debugInfo.statement).replace(/\s+/g, " ").trim() + "`");
    }
    if (debugInfo.errorLocation) parts.push(String(debugInfo.errorLocation));
  } else if (error.message) {
    parts.push(String(error.message));
  }
  return parts.length > 0 ? " (" + parts.join(" ") + ")" : "";
}

async function handleGenerateBibliography() {
  const outputElement = document.getElementById("output");
  outputElement.textContent = "Generating bibliography...";

  try {
    const citations = await collectCitationsFromSlides();
    const uniqueKeys = Array.from(new Set(citations.map((citation) => citation.key)));
    log("Found unique citation keys:", uniqueKeys);

    if (uniqueKeys.length === 0) {
      outputElement.textContent = "No citations found in the presentation.";
      return;
    }

    const style = getSelectedBibliographyStyle();
    const missingKeys = await runStep("checking the citation keys", () => findMissingCitationKeys(uniqueKeys, style));
    const bibliographyHtml = await runStep("asking Zotero", () =>
      fetchBibliographyFromServer(uniqueKeys, style, "html"),
    );
    const bibliography = parseFormattedBibliography(bibliographyHtml);

    if (!bibliography.text) {
      outputElement.textContent = "Zotero returned an empty bibliography.";
      return;
    }

    await runStep("writing the slides", () => writeBibliographySlide(bibliography, isSplitBibliographyEnabled()));

    outputElement.textContent = missingKeys.length > 0
      ? `Bibliography generated. Not found in Zotero: ${missingKeys.join(", ")}`
      : "Bibliography generated successfully!";
  } catch (error) {
    logError("Error generating bibliography:", error);
    const detail = describeError(error);
    outputElement.textContent = "Error: Could not generate bibliography." + detail;
    reportToHelper("generate bibliography failed" + detail);
    await snapshotCurrentSlide();
  }
}

/* The Zotero picker's extra fields ("see ", "p. 42", suppress author) arrive next to the item or
   inside it, so both are checked. */
function pickerField(item, name) {
  const sources = [item, item ? item.item : null];
  for (const source of sources) {
    if (!source) continue;
    const value = source[name];
    if (value !== undefined && value !== null && value !== "") {
      return String(value);
    }
  }
  return "";
}

function formatSingleCitation(item) {
  const creators = item.item.creators;
  const date = item.item.date || "";
  const year = extractYearFromDate(date);

  const suppressAuthor = ["true", "1"].includes(
    pickerField(item, "suppressAuthor") || pickerField(item, "suppress author"),
  );
  const locator = pickerField(item, "locator");
  const prefix = pickerField(item, "prefix");
  const suffix = pickerField(item, "suffix");

  let citation = year;
  if (!suppressAuthor) {
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
    citation = `${authorString}, ${year}`;
  }

  if (locator) citation += `, ${locator}`;
  if (suffix) citation += `, ${suffix}`;
  if (prefix) citation = `${prefix} ${citation}`;
  return citation;
}

function extractYearFromDate(dateStr) {
  if (!dateStr || typeof dateStr !== "string") {
    return "n.d.";
  }

  const isoMatch = /^(\d{4})[-/]/.exec(dateStr);
  if (isoMatch) {
    return isoMatch[1];
  }

  const yearMatch = /(1[5-9]\d{2}|20\d{2})/.exec(dateStr);
  if (yearMatch) {
    return yearMatch[1];
  }

  const parsed = new Date(dateStr);
  if (!isNaN(parsed.getTime())) {
    const year = parsed.getFullYear();
    if (year > 1500 && year < 2100) {
      return String(year);
    }
  }

  return "n.d.";
}

async function fetchBibliographyFromServer(keys, style, format) {
  const response = await fetch(BIB_ENDPOINT, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ keys, style, format: format || "text" }),
  });

  if (!response.ok) {
    const errorData = await response.json();
    throw new Error(errorData.error || `Server responded with status: ${response.status}`);
  }

  const data = await response.json();
  return data.bibliography;
}

async function insertCitationsIntoPowerPoint(zoteroItems) {
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.getSelectedSlides();
      slides.load("items");
      await context.sync();

      if (slides.items.length === 0) {
        logError("No slide selected.");
        document.getElementById("output").textContent = "Please select a slide first.";
        return;
      }
      const slide = slides.items[0];

      if (!Array.isArray(zoteroItems) || zoteroItems.length === 0) {
        log("No valid citation items found in the response.");
        document.getElementById("output").textContent = "No citations found to insert.";
        return;
      }

      const customTags = slide.tags;
      customTags.load("key, value");
      await context.sync();

      const existingTag = customTags.items.find((tag) => tag.key === ZOTERO_TAG_KEY);
      // {k: citation key, l: short label for the pane}; slides written earlier hold plain key strings.
      const citations = [];
      const seenKeys = new Set();
      const remember = (entry) => {
        if (!entry.key || seenKeys.has(entry.key)) return;
        seenKeys.add(entry.key);
        citations.push(entry);
      };

      if (existingTag) {
        parseCitationTag(existingTag.value).forEach(remember);
      }
      zoteroItems.forEach((item) => remember({ key: item.citationKey, label: formatSingleCitation(item) }));

      slide.tags.add(
        ZOTERO_TAG_KEY,
        JSON.stringify(citations.map((entry) => ({ k: entry.key, l: entry.label }))),
      );

      const citationParts = zoteroItems.map((item) => ({
        key: item.citationKey,
        text: formatSingleCitation(item),
      }));
      const citationsText = "(" + citationParts.map((part) => part.text).join("; ") + ")";
      const shapeTagEntries = citationParts.map((part) => ({ key: part.key, label: part.text }));
      let holderId = "";

      try {
        const selectedTextRange = context.presentation.getSelectedTextRange();
        selectedTextRange.load("text");
        await context.sync();

        const originalText = selectedTextRange.text;
        selectedTextRange.text = originalText + " " + citationsText;
        await context.sync();
      } catch (error) {
        if (error.name === "RichApi.Error" && error.code === "GeneralException") {
          log("No text range selected. Creating a new text box.");
          const textBox = slide.shapes.addTextBox(citationsText, {
            left: 100,
            top: 150,
            width: 400,
            height: 50,
          });
          textBox.load("textFrame/textRange");
          await context.sync();
        } else {
          logError("An unexpected error occurred during text insertion:", error);
          throw error;
        }
      }

      // Record the citations on the shape that received them: that is what tells the add-in later
      // whether a citation is still there, without reading the citation's wording.
      if (supportsPowerPointApi("1.3")) {
        try {
          holderId = await findCitationShapeId(context, slide, citationsText);
          if (holderId !== "") {
            // Mark on the slide which citations a shape recorded: only those can be checked when a
            // text box goes away.
            slide.tags.add(
              ZOTERO_TAG_KEY,
              JSON.stringify(
                citations.map((entry) => ({
                  k: entry.key,
                  l: entry.label,
                  s: shapeTagEntries.some((part) => part.key === entry.key) ? holderId : "",
                })),
              ),
            );
            await context.sync();
          }
        } catch (error) {
          logWarn("Could not record the citations on their shape", error);
        }
      }
    });

    await displayCitationsFromSlide();
    displayAllCitations();
  } catch (error) {
    logError("Error interacting with PowerPoint:", error);
    document.getElementById("output").textContent = "Error: Could not insert citations.";
  }
}

async function removeCitation(keyToRemove) {
  try {
    let textRemoved = false;
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const customTags = slide.tags;
      customTags.load("key, value");
      await context.sync();

      const zoteroTag = customTags.items.find((tag) => tag.key === ZOTERO_TAG_KEY);
      if (zoteroTag) {
        const citations = parseCitationTag(zoteroTag.value);
        const removedEntry = citations.find((entry) => entry.key === keyToRemove);
        const remaining = citations.filter((entry) => entry.key !== keyToRemove);
        slide.tags.add(
          ZOTERO_TAG_KEY,
          JSON.stringify(remaining.map((entry) => ({ k: entry.key, l: entry.label, s: entry.shapeId || "" }))),
        );
        await context.sync();

        // The text goes with the key, otherwise the slide still shows a citation that the
        // bibliography no longer lists - and if the text cannot be found, the key stays as well.
        if (removedEntry) {
          textRemoved = await removeCitationTextFromSlide(context, slide, removedEntry);
        }
      }
    });

    await displayCitationsFromSlide();
    displayAllCitations();
    if (!textRemoved) {
      const note = document.createElement("p");
      note.className = "note";
      note.textContent =
        "Its citation text was not found on the slide (deleted or rewritten), so only the key was removed.";
      document.getElementById("output").appendChild(note);
    }
  } catch (error) {
    logError("Error removing citation:", error);
    document.getElementById("output").textContent = "Error: Could not remove citation.";
  }
}

/* Removes a citation's text from a slide. The stored label is the text that was written for it, so
   the first shape that contains it wins; the search tolerates edited whitespace and keeps the
   other citations of a group intact. */
async function removeCitationTextFromSlide(context, slide, entry) {
  const shapes = slide.shapes;
  shapes.load("items/textFrame/textRange/text");
  await context.sync();

  for (const shape of shapes.items) {
    const textFrame = shape.textFrame;
    const textRange = textFrame && textFrame.textRange ? textFrame.textRange : null;
    if (!textRange || !textRange.text) continue;

    const span = findCitationSpan(textRange.text, entry.label || entry.key);
    const afterRemoval = span
      ? textRange.text.slice(0, span.start) + textRange.text.slice(span.end)
      : null;
    if (afterRemoval === null) continue;
    const cleaned = tidyCitationText(repairCitationText(afterRemoval, [entry.key]));
    textRange.text = cleaned;
    await context.sync();
    logInfo("removed the citation text of", entry.key);
    return true;
  }

  logWarn("no shape on this slide contained the text of", entry.key);
  return false;
}

async function displayCitationsFromSlide() {
  const outputElement = document.getElementById("output");
  outputElement.innerHTML = "";

  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      if (!slide) {
        outputElement.textContent = "Select a slide to view its citations.";
        return;
      }

      const customTags = slide.tags;
      const shapes = slide.shapes;
      customTags.load("key, value");
      shapes.load("items/textFrame/textRange/text");
      await context.sync();

      const slideText = shapes.items
        .map((shape) => {
          const textFrame = shape.textFrame;
          const textRange = textFrame && textFrame.textRange ? textFrame.textRange : null;
          return textRange && textRange.text ? textRange.text : "";
        })
        .join("\n");

      const zoteroTag = customTags.items.find((tag) => tag.key === ZOTERO_TAG_KEY);

      if (zoteroTag) {
        try {
          const citations = parseCitationTag(zoteroTag.value);
          // A citation whose text was deleted from the slide no longer belongs in the list or the
          // bibliography, so entries that are gone from the slide's text are dropped from the tag.
          let listed = citations;
          try {
            const shapesById = await readCitationShapes(context, slide);
            listed = citations.filter((citation) => !citation.recorded || shapesById.get(citation.shapeId) === true);
          } catch (error) {
            logWarn("Could not check whether the citations are still there", error);
            reportToHelper("citation shape check failed" + describeError(error));
          }
          const dropped = citations.filter((citation) => listed.indexOf(citation) < 0);
          if (dropped.length > 0) {
            const droppedKeys = dropped.map((citation) => citation.key).join(", ");
            logInfo("citations whose text is gone:", droppedKeys);
            reportToHelper("dropped citations whose text was deleted: " + droppedKeys);
            slide.tags.add(
              ZOTERO_TAG_KEY,
              JSON.stringify(listed.map((citation) => ({ k: citation.key, l: citation.label }))),
            );
            await context.sync();
          }
          for (const citation of dropped) {
            // Clean up whatever the citation left behind: an emptied marker, or text that no
            // longer belongs to the bibliography.
            await removeCitationTextFromSlide(context, slide, citation);
          }
          if (listed.length > 0) {
            const list = document.createElement("ul");

            listed.forEach((citation) => {
              const listItem = document.createElement("li");
              const removeButton = document.createElement("button");
              removeButton.className = "remove-btn";
              removeButton.innerHTML = "&times;";
              removeButton.title = `Remove citation: ${citation.key}`;
              removeButton.dataset.key = citation.key;

              const keyText = document.createElement("span");
              keyText.className = "citation-key";
              keyText.textContent = citation.label || citation.key;

              listItem.appendChild(removeButton);
              listItem.appendChild(keyText);
              list.appendChild(listItem);
            });

            outputElement.replaceChildren(list);
          } else {
            outputElement.textContent = "No Zotero citations found on this slide.";
          }
          if (dropped.length > 0) {
            const note = document.createElement("p");
            note.className = "note";
            note.textContent =
              "No longer counted (its text is gone from this slide): " +
              dropped.map((citation) => citation.key).join(", ");
            outputElement.appendChild(note);
          }
        } catch (error) {
          logError("Error parsing citation tags from slide:", error);
          reportToHelper("reading the slide citations failed" + describeError(error));
          outputElement.textContent = "Error reading citations from this slide.";
        }
      } else {
        outputElement.textContent = "No Zotero citations found on this slide.";
      }
    });
  } catch (error) {
    logError("Error displaying citations from slide:", error);
    if (outputElement.innerHTML === "") {
      outputElement.textContent = "Could not load citations for this slide.";
    }
  }
}

/* BBT is asked for HTML because that is the only way to get real italics, bold, sub- and
   superscript out of Zotero. Only that small subset of tags is interpreted. */
const BIBLIOGRAPHY_FORMAT_TAGS = {
  i: "italic",
  em: "italic",
  b: "bold",
  strong: "bold",
  sub: "subscript",
  sup: "superscript",
  sc: "smallCaps",
};

const BIBLIOGRAPHY_LINE_BREAK_TAGS = { div: true, p: true, br: true, li: true, tr: true };

const HTML_ENTITIES = {
  nbsp: " ",
  amp: "&",
  lt: "<",
  gt: ">",
  quot: "\"",
  apos: "'",
  ndash: "\u2013",
  mdash: "\u2014",
  hellip: "\u2026",
  lsquo: "\u2018",
  rsquo: "\u2019",
  ldquo: "\u201c",
  rdquo: "\u201d",
  times: "\u00d7",
};

function decodeHtmlEntities(text) {
  return text
    .replace(/&#x([0-9a-f]+);/gi, (match, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (match, code) => String.fromCodePoint(parseInt(code, 10)))
    .replace(/&([a-z]+);/gi, (match, name) =>
      Object.prototype.hasOwnProperty.call(HTML_ENTITIES, name.toLowerCase())
        ? HTML_ENTITIES[name.toLowerCase()]
        : match,
    );
}

/* Turns the HTML bibliography into plain text plus the formatting runs that index into it,
   where the formatting is the flag set that must be applied to that slice of text. */
function parseFormattedBibliography(html) {
  const runs = [];
  const openFlags = [];
  let text = "";
  let pending = "";
  let pendingFormat = null;

  const flush = () => {
    if (pending.length > 0 && pendingFormat) {
      runs.push({ start: text.length - pending.length, length: pending.length, format: pendingFormat });
    }
    pending = "";
  };

  const append = (chunk, format) => {
    if (chunk.length === 0) return;
    const formatKey = format ? JSON.stringify(format) : "";
    const pendingKey = pendingFormat ? JSON.stringify(pendingFormat) : "";
    if (formatKey !== pendingKey) {
      flush();
      pendingFormat = format;
    }
    text += chunk;
    pending += chunk;
  };

  // One line break between entries; the whitespace between block tags must not become a blank line.
  const appendLineBreak = () => {
    if (text.length > 0 && !text.endsWith("\n")) append("\n", null);
  };

  const appendText = (chunk, format) => {
    const collapsed = chunk.replace(/[ \t\r\n\f\v]+/g, (run) => (run.indexOf("\n") >= 0 ? "\n" : " "));
    collapsed.split("\n").forEach((piece, index) => {
      if (index > 0) appendLineBreak();
      if (piece.length > 0) append(piece, format);
    });
  };

  const tokens = String(html || "").match(/<[^>]*>|[^<]+/g) || [];
  for (const token of tokens) {
    if (token.charAt(0) === "<") {
      const tag = /^<\/?\s*([a-zA-Z0-9]+)/.exec(token);
      if (!tag) continue;
      const name = tag[1].toLowerCase();
      const closing = token.indexOf("</") === 0;

      if (BIBLIOGRAPHY_LINE_BREAK_TAGS[name]) {
        if (closing || name === "br") {
          flush();
          appendLineBreak();
        }
        continue;
      }

      if (name === "span") {
        // Only small caps matters; </span> closes the matching <span>, whatever it pushed.
        if (closing) {
          const spanIndex = openFlags.lastIndexOf("smallCaps");
          if (spanIndex >= 0) openFlags.splice(spanIndex, 1);
        } else if (/small-caps/i.test(token)) {
          openFlags.push("smallCaps");
        }
        continue;
      }

      const flag = BIBLIOGRAPHY_FORMAT_TAGS[name];
      if (!flag) continue;

      if (closing) {
        const index = openFlags.lastIndexOf(flag);
        if (index >= 0) openFlags.splice(index, 1);
      } else if (token.indexOf("/>") !== token.length - 2) {
        openFlags.push(flag);
      }
      continue;
    }

    const format = {};
    for (const flag of openFlags) {
      format[flag] = true;
    }
    appendText(decodeHtmlEntities(token), openFlags.length > 0 ? format : null);
  }
  flush();

  return withBibliographyEntries(trimFormattedText(text, runs));
}

/* Splits the bibliography into entries (one per line) so it can be spread over several slides.
   Each entry keeps its own formatting runs, offset from its own start. */
function splitIntoBibliographyEntries(bibliography) {
  const text = bibliography.text;
  const entries = [];
  let start = 0;

  for (let end = 0; end <= text.length; end++) {
    if (end < text.length && text.charAt(end) !== "\n") continue;
    const entryText = text.slice(start, end);
    if (entryText.length > 0) {
      const entryRuns = [];
      for (const run of bibliography.runs) {
        const runStart = Math.max(run.start, start);
        const runStop = Math.min(run.start + run.length, end);
        if (runStop > runStart) {
          entryRuns.push({ start: runStart - start, length: runStop - runStart, format: run.format });
        }
      }
      entries.push({ text: entryText, runs: entryRuns });
    }
    start = end + 1;
  }
  return entries;
}

function withBibliographyEntries(bibliography) {
  bibliography.entries = splitIntoBibliographyEntries(bibliography);
  return bibliography;
}

/* The inverse: one page's worth of entries back into text plus runs. */
function joinBibliographyEntries(entries) {
  let text = "";
  const runs = [];
  entries.forEach((entry, index) => {
    if (index > 0) text += "\n";
    const offset = text.length;
    text += entry.text;
    entry.runs.forEach((run) => {
      runs.push({ start: run.start + offset, length: run.length, format: run.format });
    });
  });
  return { text: text, runs: runs };
}

/* Drops surrounding whitespace while keeping the run offsets pointing at the same characters. */
function trimFormattedText(text, runs) {
  const leading = text.length - text.replace(/^\s+/, "").length;
  const trimmed = text.replace(/^\s+/, "").replace(/\s+$/, "");
  const end = leading + trimmed.length;
  const adjusted = [];
  for (const run of runs) {
    const start = Math.max(run.start, leading);
    const stop = Math.min(run.start + run.length, end);
    if (stop > start) {
      adjusted.push({ start: start - leading, length: stop - start, format: run.format });
    }
  }
  return { text: trimmed, runs: adjusted };
}

function applyBibliographyFormatting(textRange, runs) {
  for (const run of runs) {
    const substring = textRange.getSubstring(run.start, run.length);
    if (!substring) continue;
    const font = substring.font;
    if (run.format.italic) font.italic = true;
    if (run.format.bold) font.bold = true;
    if (run.format.subscript) font.subscript = true;
    if (run.format.superscript) font.superscript = true;
    if (run.format.smallCaps) font.smallCaps = true;
  }
}

async function addBibliographySlideFromMaster(context, slides) {
  const slideMasters = context.presentation.slideMasters.load("id, name, layouts/items/name, layouts/items/id");
  await context.sync();

  let targetMaster = slideMasters.items[0] || null;
  let layoutId = null;

  outer: for (const master of slideMasters.items) {
    for (const layout of master.layouts.items) {
      const layoutName = (layout.name || "").toLowerCase();
      if (layoutName === "title and content" || (layoutName.includes("title") && layoutName.includes("content"))) {
        targetMaster = master;
        layoutId = layout.id;
        break outer;
      }
    }
  }

  if (!targetMaster) {
    throw new Error("No slide master available for bibliography slide creation.");
  }
  if (!layoutId) {
    layoutId = (targetMaster.layouts.items[1] && targetMaster.layouts.items[1].id) || null;
  }
  if (!layoutId && targetMaster.layouts.items[0]) {
    layoutId = targetMaster.layouts.items[0].id;
  }

  const options = { slideMasterId: targetMaster.id };
  if (layoutId) options.layoutId = layoutId;
  slides.add(options);
  await context.sync();

  const count = slides.getCount();
  await context.sync();
  slides.load("items");
  await context.sync();
  return slides.getItemAt(count.value - 1);
}

/* A slide whose title already says "References" is reused, so decks made before the tag existed
   do not collect a second References slide. The slides we tagged are found by their tag instead. */
async function findBibliographySlideByTitle(context, slides) {
  // The shape text must be loaded and synced before it can be read below.
  try {
    for (const slide of slides.items) {
      slide.shapes.load("items/name,items/textFrame/textRange/text");
    }
    await context.sync();
  } catch (error) {
    logWarn("Could not look for an existing References slide", error);
    return null;
  }

  const byTitle = slides.items.find((slide) =>
    slide.shapes.items.some((shape) => {
      const textFrame = shape.textFrame;
      const shapeText = textFrame && textFrame.textRange ? textFrame.textRange.text || "" : "";
      return shapeText.trim().toLowerCase() === BIBLIOGRAPHY_TITLE.toLowerCase();
    }),
  );
  return byTitle || null;
}


/* Finds (and if asked, creates) the title and content shapes of a slide. The shapes must be
   loaded and synced before the names can be read. */
async function findSlideTextShapes(context, slide, createMissing) {
  slide.shapes.load("items/name");
  await context.sync();

  let titleShape = null;
  let contentShape = null;
  for (const shape of slide.shapes.items) {
    const name = (shape.name || "").toLowerCase();
    if (!titleShape && name.includes("title")) {
      titleShape = shape;
    } else if (!contentShape && (name.includes("content") || name.includes("body"))) {
      contentShape = shape;
    }
  }
  if (!titleShape && slide.shapes.items[0]) titleShape = slide.shapes.items[0];
  if (!contentShape && slide.shapes.items[1]) contentShape = slide.shapes.items[1];

  if (createMissing && !titleShape) {
    titleShape = slide.shapes.addTextBox("", { left: 50, top: 50, width: 860, height: 100 });
  }
  if (createMissing && !contentShape) {
    contentShape = slide.shapes.addTextBox("", { left: 50, top: 150, width: 860, height: 350 });
  }
  return { titleShape: titleShape, contentShape: contentShape };
}

async function fillBibliographySlide(context, slide, page, title) {
  const shapes = await findSlideTextShapes(context, slide, true);
  shapes.titleShape.textFrame.textRange.text = title;
  shapes.titleShape.textFrame.textRange.font.size = 44;

  const textRange = shapes.contentShape.textFrame.textRange;
  textRange.text = page.text;
  textRange.font.size = getBibliographyFontSize();
  applyBibliographyFormatting(textRange, page.runs);
  shapes.contentShape.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeShapeToFitText;
  await context.sync();
}

/* Spreads the entries over as many slides as they need. The measurement is PowerPoint's own: with
   autoSizeShapeToFitText the shape grows to the text, so its height says how much room that text
   wants. The height the content placeholder starts with is the space one slide offers. */
async function splitBibliographyIntoPages(context, slide, bibliography, splitAcrossSlides) {
  const entries = bibliography.entries || [];
  if (!splitAcrossSlides || entries.length <= 1) {
    return [joinBibliographyEntries(entries)];
  }

  try {
    return await measureBibliographyPages(context, slide, entries);
  } catch (error) {
    logWarn("Could not measure the bibliography on the slide:", error);
    reportToHelper("bibliography measurement failed" + describeError(error));
    return splitEntriesByCharacterBudget(entries, BIBLIOGRAPHY_CHARS_PER_SLIDE);
  }
}

/* Measures how much of the bibliography fits on one slide by putting the text into a scratch text
   box of the same width (PowerPoint grows a text box to fit its text) and comparing that height
   with the room one slide offers. Two things are deliberately not trusted: the content placeholder
   on the slide, which may have grown to fit an earlier, longer bibliography, and a height that does
   not react to the text at all - in either case the measurement is rejected and the caller falls
   back to a character budget. */
async function measureBibliographyPages(context, slide, entries) {
  const shapes = await findSlideTextShapes(context, slide, true);
  const contentShape = shapes.contentShape;
  contentShape.load("name,height,width");
  await context.sync();

  const budget = await slideTextBudget(context, slide, contentShape);
  const ruler = slide.shapes.addTextBox("", { left: -3000, top: 0, width: contentShape.width, height: 40 });
  ruler.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeShapeToFitText;
  ruler.textFrame.textRange.font.size = getBibliographyFontSize();

  try {
    const measure = async (candidate) => {
      ruler.textFrame.textRange.text = joinBibliographyEntries(candidate).text;
      ruler.load("height");
      slide.shapes.load("items/height");   // loading one property twice may not refresh it
      await context.sync();
      return ruler.height;
    };

    const allHeight = await measure(entries);
    const oneHeight = await measure(entries.slice(0, 1));
    const measured = "one=" + oneHeight + " all=" + allHeight + " budget=" + budget;
    logInfo("bibliography text height:", measured);
    reportToHelper("bibliography text height: " + measured);
    if (!(oneHeight < allHeight)) {
      throw new Error("the measured text height does not react to the text (" + measured + ")");
    }

    const pages = [];
    let start = 0;
    while (start < entries.length) {
      const remaining = entries.slice(start);
      let fits = 1;
      if ((await measure(remaining)) <= budget + 1) {
        fits = remaining.length;
      } else {
        let low = 1;
        let high = remaining.length;
        while (low <= high) {
          const middle = Math.floor((low + high) / 2);
          if ((await measure(remaining.slice(0, middle))) <= budget + 1) {
            fits = Math.max(fits, middle);
            low = middle + 1;
          } else {
            high = middle - 1;
          }
        }
      }
      logInfo("bibliography page holds entries", start, "to", start + fits - 1);
      pages.push(joinBibliographyEntries(remaining.slice(0, fits)));
      start += fits;
    }
    return pages;
  } finally {
    ruler.delete();
    await context.sync();
  }
}

/* The layout keeps the height the placeholder was designed with, which is the room a slide offers;
   the slide's own placeholder may have grown to fit a previous bibliography. */
async function slideTextBudget(context, slide, contentShape) {
  const layoutShapes = slide.layout.shapes;
  layoutShapes.load("items/name,items/height");
  await context.sync();
  const match = layoutShapes.items.find((shape) => (shape.name || "") === (contentShape.name || ""));
  if (!match || !(match.height > 0)) {
    throw new Error("no layout placeholder matches " + contentShape.name);
  }
  return match.height;
}

/* The fallback splitter: no PowerPoint measurement, just a character budget per slide. Entries are
   never split, so an entry longer than the budget gets a page of its own. */
function splitEntriesByCharacterBudget(entries, budget) {
  const pages = [];
  let current = [];
  let length = 0;

  for (const entry of entries) {
    const projected = current.length === 0 ? entry.text.length : length + 1 + entry.text.length;
    if (current.length > 0 && projected > budget) {
      pages.push(joinBibliographyEntries(current));
      current = [];
      length = 0;
    }
    current.push(entry);
    length = current.length === 1 ? entry.text.length : length + 1 + entry.text.length;
  }
  if (current.length > 0) pages.push(joinBibliographyEntries(current));
  return pages;
}

/* Writes the bibliography, spread over as many slides as it needs, into the slides we created for
   it before (their page number is the tag value). Slides from an earlier, longer run are removed. */
async function writeBibliographySlide(bibliography, splitAcrossSlides) {
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    for (const slide of slides.items) {
      slide.tags.load("key, value");
    }
    await context.sync();

    const ours = slides.items
      .map((slide) => ({ slide: slide, tag: slide.tags.items.find((tag) => tag.key === ZOTERO_BIBLIOGRAPHY_TAG) }))
      .filter((entry) => entry.tag);
    const pageSlide = (page) => {
      const found = ours.find((entry) => entry.tag.value === String(page));
      return found ? found.slide : null;
    };

    const first =
      pageSlide(1) ||
      (await findBibliographySlideByTitle(context, slides)) ||
      (await addBibliographySlideFromMaster(context, slides));
    const pages = await splitBibliographyIntoPages(context, first, bibliography, splitAcrossSlides);

    const targets = [first];
    for (let page = 2; page <= pages.length; page++) {
      targets.push(pageSlide(page) || (await addBibliographySlideFromMaster(context, slides)));
    }

    for (const entry of ours) {
      if (!targets.includes(entry.slide)) {
        logInfo("removing a bibliography slide that is no longer needed");
        entry.slide.delete();
      }
    }

    for (let index = 0; index < targets.length; index++) {
      targets[index].tags.add(ZOTERO_BIBLIOGRAPHY_TAG, String(index + 1));
    }
    await context.sync();

    for (let index = 0; index < targets.length; index++) {
      const title = index === 0 ? BIBLIOGRAPHY_TITLE : BIBLIOGRAPHY_CONTINUATION_TITLE;
      await fillBibliographySlide(context, targets[index], pages[index], title);
      await snapshotSlide(context, targets[index], "references-page-" + (index + 1));
    }

    // The bibliography belongs at the end of the deck, first page first - but Slide.moveTo throws a
    // GeneralException on PowerPoint builds that do not support it (16.0.17932 on Windows does not),
    // so it is attempted once and skipped afterwards instead of failing the whole run.
    slides.load("items");
    await context.sync();
    if (canMoveSlides && targets.length > 0) {
      const firstPosition = slides.items.length - targets.length + 1;
      try {
        for (let index = 0; index < targets.length; index++) {
          if (firstPosition + index >= 1) targets[index].moveTo(firstPosition + index);
        }
        await context.sync();
      } catch (error) {
        canMoveSlides = false;
        logWarn("This PowerPoint build cannot move slides; the bibliography stays where it is", error);
        reportToHelper("slide.moveTo is not supported on this build" + describeError(error));
      }
    }
  });
}
