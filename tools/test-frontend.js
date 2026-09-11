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
    document: { readyState: "complete", getElementById: () => null, querySelector: () => null },
    window: {
        localStorage: { getItem: () => null, setItem: () => {} },
        sessionStorage: { getItem: () => null, setItem: () => {} },
        location: { reload: () => {} },
    },
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

console.log("bibliography entries");
const twoEntries = sandbox.parseFormattedBibliography(
    '<div class="csl-entry">Smith, J. (2020). <i>Title one</i>.</div>\n  <div class="csl-entry">Doe, A. (2019). Title two.</div>',
);
check(
    "whitespace between entries does not become an empty line",
    twoEntries.text === "Smith, J. (2020). Title one.\nDoe, A. (2019). Title two.",
    JSON.stringify(twoEntries.text),
);
check("two entries", twoEntries.entries.length === 2, JSON.stringify(twoEntries.entries.map((entry) => entry.text)));
check(
    "an entry keeps its own formatting",
    twoEntries.entries[0].runs.length === 1 &&
        twoEntries.entries[0].text.slice(
            twoEntries.entries[0].runs[0].start,
            twoEntries.entries[0].runs[0].start + twoEntries.entries[0].runs[0].length,
        ) === "Title one",
    JSON.stringify(twoEntries.entries[0]),
);
check("the unformatted entry has no runs", twoEntries.entries[1].runs.length === 0);
check(
    "a page keeps the formatting of its entries",
    twoEntries.entries.every((entry) => {
        const page = sandbox.joinBibliographyEntries([entry]);
        return (
            page.text === entry.text &&
            page.runs.every((run) => page.text.slice(run.start, run.start + run.length) === entry.text.slice(run.start, run.start + run.length))
        );
    }),
);
check(
    "joining entries rebases the run offsets",
    (() => {
        const page = sandbox.joinBibliographyEntries(twoEntries.entries);
        const run = page.runs[0];
        return run.start > 0 && page.text.slice(run.start, run.start + run.length) === "Title one";
    })(),
);
check(
    "pretty printed entries still give one line each",
    sandbox.parseFormattedBibliography('<div class="csl-entry">\n  First.\n</div>\n<div class="csl-entry">\n  Second.\n</div>').text === "First.\nSecond.",
);

console.log("page splitting without PowerPoint");
const budgetEntries = [1, 2, 3, 4, 5].map((index) => ({ text: "entry " + index, runs: [] }));
const budgetPages = sandbox.splitEntriesByCharacterBudget(budgetEntries, 16);
check(
    "a page holds as many entries as the budget allows",
    budgetPages.length === 3,
    JSON.stringify(budgetPages.map((page) => page.text)),
);
check("the first page stops before the budget", budgetPages[0].text === "entry 1\nentry 2", JSON.stringify(budgetPages[0].text));
check("no entry is lost or duplicated", budgetPages.map((page) => page.text).join("\n").split("\n").length === 5);
check(
    "an entry bigger than the budget gets its own page",
    sandbox.splitEntriesByCharacterBudget(
        [{ text: "small", runs: [] }, { text: "x".repeat(90), runs: [] }, { text: "small", runs: [] }],
        16,
    ).length === 3,
);
check(
    "page splitting keeps formatting runs valid",
    (() => {
        const pages = sandbox.splitEntriesByCharacterBudget(twoEntries.entries, 20);
        return pages.every((page) =>
            page.runs.every((run) => page.text.slice(run.start, run.start + run.length) === "Title one"),
        );
    })(),
);

console.log("citation tags");
check(
    "the old key-only tag still parses",
    JSON.stringify(sandbox.parseCitationTag('["smith2020","doe2019"]')) === JSON.stringify([
        { key: "smith2020", label: "" },
        { key: "doe2019", label: "" },
    ]),
    JSON.stringify(sandbox.parseCitationTag('["smith2020","doe2019"]')),
);
check(
    "the label is read back",
    JSON.stringify(sandbox.parseCitationTag('[{"k":"smith2020","l":"Smith, 2020"}]')) === JSON.stringify([
        { key: "smith2020", label: "Smith, 2020" },
    ]),
);
check("broken tags are ignored", sandbox.parseCitationTag("not json").length === 0);
check("unknown shapes are ignored", sandbox.parseCitationTag('[null,3,{"x":1}]').length === 0);
check(
    "citations are grouped by key with their slides",
    (() => {
        const groups = sandbox.groupCitationsByKey([
            { key: "a", label: "(A, 2020)", slideNumber: 2 },
            { key: "b", label: "", slideNumber: 3 },
            { key: "a", label: "", slideNumber: 5 },
            { key: "a", label: "(A, 2020)", slideNumber: 5 },
        ]);
        return (
            groups.length === 2 &&
            JSON.stringify(groups[0]) === JSON.stringify({ key: "a", label: "(A, 2020)", slides: [2, 5] }) &&
            JSON.stringify(groups[1]) === JSON.stringify({ key: "b", label: "", slides: [3] })
        );
    })(),
    JSON.stringify(sandbox.groupCitationsByKey([{ key: "a", label: "", slideNumber: 1 }])),
);

/* --- Office.js usage lint ----------------------------------------------------------------
 * Office.js proxies expose nothing until a property has been queued with load() and delivered by
 * an awaited context.sync(). Reading too early throws "The property 'x' is not available" at run
 * time - inside PowerPoint, where that is slow to discover. The rules below catch that class of
 * bug here instead, and the fixtures prove the rules actually fire (a lint that cannot flag the
 * bug it was written for is worse than no lint).
 */
const OFFICEJS_LINT_FIXTURES = [
    {
        name: "collection read after load without a sync (the real regression)",
        source: [
            "async function findBibliographySlide(slides) {",
            "  for (const slide of slides.items) {",
            "    slide.shapes.load(\"items/name\");",
            "  }",
            "  return slides.items.find((slide) => slide.shapes.items.length > 0);",
            "}",
        ].join("\n"),
        expect: "slide.shapes.items",
    },
    {
        name: "sync not awaited",
        source: [
            "async function loadThem(context) {",
            "  const slides = context.presentation.slides;",
            "  slides.load(\"items\");",
            "  context.sync();",
            "  return slides.items.length;",
            "}",
        ].join("\n"),
        expect: "awaited",
    },
    {
        name: "getCount() value read before the sync",
        source: [
            "async function countThem(slides) {",
            "  const count = slides.getCount();",
            "  return count.value;",
            "}",
        ].join("\n"),
        expect: "count.value",
    },
    {
        name: "nested load makes the child collection readable",
        source: [
            "async function ok(context) {",
            "  const masters = context.presentation.slideMasters.load(\"id, layouts/items/name\");",
            "  await context.sync();",
            "  for (const master of masters.items) {",
            "    for (const layout of master.layouts.items) {",
            "      void layout.name;",
            "    }",
            "  }",
            "}",
        ].join("\n"),
        expect: null,
    },
];

/* Line ranges of the function bodies, found by indentation: this file keeps one statement per
   line, so the body ends at the first closing brace back at the head's indentation. */
/* Function body line ranges, found by indentation: one statement per line, so a body ends at the
   first closing brace back at the head's own indentation. Heads may sit mid-line (e.g.
   `await PowerPoint.run(async (context) => {`), which is why the indentation comes from the line. */
function officeJsScopes(lines) {
    const headPatterns = [
        /(?:async\s+)?function\s+[\w$]*\s*\(([^)]*)\)\s*\{/,
        /(?:async\s+)?\(([^()]*)\)\s*=>\s*\{/,
        /([\w$]+)\s*=>\s*\{/,
    ];
    const scopes = [];
    for (let i = 0; i < lines.length; i++) {
        let params = null;
        for (const pattern of headPatterns) {
            const match = pattern.exec(lines[i]);
            if (match) {
                params = match[1] || "";
                break;
            }
        }
        if (params === null) continue;

        const indent = lines[i].search(/\S/);
        let end = lines.length - 1;
        for (let j = i + 1; j < lines.length; j++) {
            const lineIndent = lines[j].search(/\S/);
            if (lineIndent >= 0 && /^\s*\}/.test(lines[j]) && lineIndent <= indent) {
                end = j;
                break;
            }
        }
        scopes.push({
            start: i,
            end,
            params: params
                .split(",")
                .map((param) => param.trim().replace(/^[\.{]*/, "").split("=")[0].trim())
                .filter(Boolean),
        });
    }
    return scopes;
}

function lintOfficeJs(source) {
    const lines = source.split("\n");
    const problems = [];
    const seen = new Set();
    const loadPattern = /(?:([\w$]+)\s*=\s*)?([\w$.\[\]]+)\.load\(\s*(['\"])([^'\"]*)\3/g;

    const report = (lineIndex, message) => {
        const key = lineIndex + ":" + message;
        if (seen.has(key)) return;
        seen.add(key);
        problems.push({ line: lineIndex + 1, text: lines[lineIndex].trim(), message });
    };

    const tokensOf = (load) => {
        const tokens = [load[2]];
        if (load[1]) tokens.push(load[1]);
        // "layouts/items/name" also makes master.layouts.items readable.
        (load[4].match(/([\w$]+)\/items\b/g) || []).forEach((segment) => tokens.push(segment.split("/")[0] + ".items"));
        return tokens;
    };

    // A helper may read a collection that its caller loaded, so a token loaded earlier in the file
    // counts as loaded; the ordering rule that actually catches bugs is the per-scope one below.
    const firstLoad = new Map();
    lines.forEach((line, index) => {
        for (const load of line.replace(/\/\/.*$/, "").matchAll(loadPattern)) {
            for (const token of tokensOf(load)) {
                if (!firstLoad.has(token)) firstLoad.set(token, index);
            }
        }
    });

    for (const scope of officeJsScopes(lines)) {
        const pending = new Set();        // loaded, but no awaited sync has delivered it yet
        const readable = new Set();       // loaded and synced
        const pendingValues = new Set();  // getCount() results that are not synced yet

        // The head line itself belongs to the enclosing scope (it may only be a head part-way
        // through, e.g. `slide.shapes.items.some((shape) => {`).
        for (let i = scope.start + 1; i <= scope.end; i++) {
            const code = lines[i].replace(/\/\/.*$/, "");
            if (!/\S/.test(code)) continue;

            const syncCall = /context\.sync\(\)/.test(code);
            if (syncCall && !/await\s+context\.sync\(\)/.test(code)) {
                report(i, "context.sync() must be awaited, otherwise nothing is loaded yet");
            }

            // A load also makes the variable it is assigned to refer to that collection.
            const loadedTargets = new Set();
            for (const load of code.matchAll(loadPattern)) {
                loadedTargets.add(load[2]);
                for (const token of tokensOf(load)) pending.add(token);
            }

            const count = /(?:const|let|var)\s+([\w$]+)\s*=\s*[\w$.\[\]]+\.getCount\(\s*\)/.exec(code);
            if (count) pendingValues.add(count[1]);

            const readPattern = /([\w$.\[\]]+)\.items\b/g;
            for (const read of code.matchAll(readPattern)) {
                const target = read[1];
                const root = target.split(/[.[]/)[0];
                if (scope.params.includes(root) || loadedTargets.has(target)) continue;
                const tokens = [target, target.split(".").slice(-1)[0] + ".items"];
                if (tokens.some((token) => pending.has(token))) {
                    report(i, target + ".items is read after load() but before an awaited context.sync()");
                } else if (!tokens.some((token) => readable.has(token)) && tokens.every((token) => !firstLoad.has(token) || firstLoad.get(token) > i)) {
                    report(i, target + ".items is read without load()");
                }
            }

            const value = /\b([\w$]+)\.value\b/.exec(code);
            if (value && pendingValues.has(value[1])) {
                report(i, value[1] + ".value is read before the getCount() request was synced");
            }

            if (syncCall) {
                pending.forEach((token) => readable.add(token));
                pending.clear();
                pendingValues.clear();
            }
        }
    }
    return problems;
}

console.log("Office.js usage lint");
OFFICEJS_LINT_FIXTURES.forEach((fixture) => {
    const found = lintOfficeJs(fixture.source);
    if (fixture.expect === null) {
        check("fixture stays clean: " + fixture.name, found.length === 0, JSON.stringify(found));
    } else {
        check(
            "fixture is flagged: " + fixture.name,
            found.some((problem) => (problem.message + " " + problem.text).includes(fixture.expect)),
            JSON.stringify(found),
        );
    }
});

["shared/frontend_core.js", "zotero-addon/www/frontend_core.js"].forEach((relative) => {
    const file = path.join(__dirname, "..", relative);
    if (!fs.existsSync(file)) return;
    const problems = lintOfficeJs(fs.readFileSync(file, "utf8"));
    check(
        "no Office.js load/sync misuse in " + relative,
        problems.length === 0,
        problems.map((problem) => problem.line + ": " + problem.message).join("; "),
    );
});

console.log("");
console.log("Passed: " + passed + "   Failed: " + failures.length);
if (failures.length > 0) {
    console.log("Failures:");
    failures.forEach((failure) => console.log("  - " + failure));
    process.exit(1);
}
