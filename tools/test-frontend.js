#!/usr/bin/env node
/*
 * Checks the pure helpers in shared/frontend_core.js without a browser or PowerPoint:
 * the citation text builder (including the picker's locator/prefix/suffix fields) and the
 * HTML bibliography parser that turns BBT's HTML into plain text plus formatting runs.
 *
 *   node tools/test-frontend.js
 */
"use strict";

const fs = require("fs");
const path = require("path");
const vm = require("vm");

const sourcePath = path.join(__dirname, "..", "shared", "frontend_core.js");
const source = fs.readFileSync(sourcePath, "utf8");

// frontend_core.js talks to Office at load time, so it runs in a sandbox with a stub host.
const sandbox = {
    Office: { onReady() {} },
    console,
    JSON,
    Math,
    String,
    Number,
    Array,
    Object,
    Set,
    Promise,
    RegExp,
    Date,
    parseInt,
    parseFloat,
    isNaN,
    document: { readyState: "complete", getElementById: () => null },
    window: {},
};
vm.createContext(sandbox);
vm.runInContext(source, sandbox, { filename: "frontend_core.js" });

let passed = 0;
const failures = [];

function check(name, condition, detail) {
    if (condition) {
        passed += 1;
        console.log("  ok   " + name);
    } else {
        failures.push(name + (detail ? " -> " + detail : ""));
        console.log("  FAIL " + name + (detail ? " -> " + detail : ""));
    }
}

function item(extra) {
    const base = {
        citationKey: "smith2020",
        item: { creators: [{ lastName: "Smith" }], date: "2020" },
    };
    return Object.assign(base, extra);
}

function runSummary(text, runs) {
    return runs.map((run) => text.slice(run.start, run.start + run.length) + "[" + Object.keys(run.format).join("+") + "]").join(" | ");
}

console.log("citation text");
check("plain", sandbox.formatSingleCitation(item()) === "Smith, 2020", sandbox.formatSingleCitation(item()));
check("locator", sandbox.formatSingleCitation(item({ locator: "p. 42" })) === "Smith, 2020, p. 42");
check("prefix", sandbox.formatSingleCitation(item({ prefix: "see" })) === "see Smith, 2020");
check("suffix", sandbox.formatSingleCitation(item({ suffix: "cf. Doe" })) === "Smith, 2020, cf. Doe");
check(
    "prefix + locator + suffix",
    sandbox.formatSingleCitation(item({ prefix: "see", locator: "p. 42", suffix: "and others" })) ===
        "see Smith, 2020, p. 42, and others",
);
check("suppress author", sandbox.formatSingleCitation(item({ suppressAuthor: true })) === "2020");
check(
    "fields inside item",
    sandbox.formatSingleCitation({ citationKey: "x", item: { creators: [{ lastName: "Doe" }], date: "2019", locator: "p. 7" } }) ===
        "Doe, 2019, p. 7",
);
check(
    "two creators",
    sandbox.formatSingleCitation({ citationKey: "x", item: { creators: [{ lastName: "A" }, { lastName: "B" }], date: "2001" } }) ===
        "A & B, 2001",
);
check(
    "three creators",
    sandbox.formatSingleCitation({
        citationKey: "x",
        item: { creators: [{ lastName: "A" }, { lastName: "B" }, { lastName: "C" }], date: "2001" },
    }) === "A et al., 2001",
);
check("no date", sandbox.formatSingleCitation({ citationKey: "x", item: { creators: [{ lastName: "A" }] } }) === "A, n.d.");

console.log("bibliography parser");
const html =
    '<div class="csl-entry">Smith, J. (2020). Title. <i>Journal Name</i>, <b>12</b>(3), 45&#8211;67.</div>' +
    '<div class="csl-entry">Doe, A. (2019). <span style="font-variant:small-caps;">agu</span> H<sub>2</sub>O m<sup>3</sup>.</div>';
const parsed = sandbox.parseFormattedBibliography(html);
const summary = runSummary(parsed.text, parsed.runs);

check("one entry per line", parsed.text.split("\n").length === 2, JSON.stringify(parsed.text));
check("tags removed, entities decoded", parsed.text.indexOf("<") < 0 && parsed.text.indexOf("45\u201367") >= 0, parsed.text);
check("italic run", summary.indexOf("Journal Name[italic]") >= 0, summary);
check("bold run", summary.indexOf("12[bold]") >= 0, summary);
check("small caps run", summary.indexOf("agu[smallCaps]") >= 0, summary);
check("subscript and superscript runs", summary.indexOf("2[subscript]") >= 0 && summary.indexOf("3[superscript]") >= 0, summary);
check(
    "every run addresses real text",
    parsed.runs.every((run) => run.length > 0 && run.start >= 0 && run.start + run.length <= parsed.text.length),
    summary,
);
check("no surrounding whitespace", parsed.text === parsed.text.trim(), JSON.stringify(parsed.text));
check("plain text produces no runs", sandbox.parseFormattedBibliography("Smith, J. (2020). Title.").runs.length === 0);
check("empty input is empty", sandbox.parseFormattedBibliography("").text === "");
check(
    "unclosed tags do not break the text",
    sandbox.parseFormattedBibliography("<i>Half a title").text === "Half a title",
);

console.log("");
console.log("Passed: " + passed + "   Failed: " + failures.length);
if (failures.length > 0) {
    console.log("Failures:");
    failures.forEach((failure) => console.log("  - " + failure));
    process.exit(1);
}
