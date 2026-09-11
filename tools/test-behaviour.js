/* Behaviour tests: the pane's functions are run against a small PowerPoint model (slides, shapes,
   text, tags) so that the flows that broke in practice - inserting, editing, deleting, removing -
   are checked end to end instead of one helper at a time. Office.js's load/sync bookkeeping is not
   modelled (the lint in tools/test-frontend.js checks that); here `load()` is a no-op and `sync()`
   resolves, which is enough to see what the add-in writes.

   Run: node tools/test-behaviour.js
*/
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const source = fs.readFileSync(path.join(__dirname, "..", "shared", "frontend_core.js"), "utf8");

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

/* --- a very small PowerPoint ------------------------------------------------------------- */

function tagList(initial) {
    const items = (initial || []).map((entry) => ({ key: entry[0], value: entry[1] }));
    return {
        items: items,
        load() {},
        add(key, value) {
            const found = items.find((tag) => tag.key === key);
            if (found) found.value = value;
            else items.push({ key: key, value: value });
        },
    };
}

function shape(name, text) {
    const range = { text: text, load() {}, font: {} };
    return {
        name: name,
        textFrame: { textRange: range, autoSizeSetting: 0 },
        tags: tagList(),
    };
}

function deck() {
    const slide = {
        tags: tagList(),
        shapes: {
            items: [],
            load() {},
            addTextBox(text) {
                const created = shape("TextBox " + (slide.shapes.items.length + 1), text);
                slide.shapes.items.push(created);
                return created;
            },
        },
    };
    const selected = shape("Content Placeholder", "");
    slide.shapes.items.push(selected);

    const context = {
        presentation: {
            slides: { items: [slide], load() {} },
            getSelectedSlides() {
                return { items: [slide], load() {}, getItemAt: () => slide };
            },
            getSelectedTextRange() {
                return selected.textFrame.textRange;
            },
        },
        sync() {
            return Promise.resolve();
        },
    };

    return {
        slide: slide,
        selected: selected,
        text: () => selected.textFrame.textRange.text,
        setText: (value) => {
            selected.textFrame.textRange.text = value;
        },
        slideTag: () => {
            const tag = slide.tags.items.find((item) => item.key === "ZOTERO_CITATION_KEYS");
            return tag ? JSON.parse(tag.value) : [];
        },
        shapeTag: (target) => {
            const tag = target.tags.items.find((item) => item.key === "ZOTERO_CITATION_KEYS");
            return tag ? JSON.parse(tag.value) : [];
        },
        /* state for the sandbox */
        PowerPoint: {
            run: async (callback) => callback(context),
            ShapeAutoSize: { autoSizeShapeToFitText: 0 },
        },
    };
}

function element() {
    const node = {
        textContent: "",
        innerHTML: "",
        className: "",
        dataset: {},
        children: [],
        appendChild(child) {
            node.children.push(child);
            return child;
        },
        replaceChildren(...children) {
            node.children = children;
        },
        addEventListener() {},
    };
    return node;
}

function loadPane(world) {
    const output = element();
    const sandbox = {
        Office: {
            onReady() {},
            context: { requirements: { isSetSupported: () => true }, diagnostics: { host: "test" } },
        },
        PowerPoint: world.PowerPoint,
        console: { log() {}, warn() {}, error(...args) { process.stderr.write("[pane error] " + args.map(String).join(" ") + "\n"); }, info() {} },
        JSON,
        Math,
        String,
        Number,
        Array,
        Object,
        Set,
        Map,
        Promise,
        RegExp,
        isNaN,
        fetch: () => Promise.resolve({ ok: true, json: () => Promise.resolve({}) }),
        document: {
            readyState: "complete",
            getElementById: () => output,
            querySelector: () => null,
            createElement: () => element(),
        },
        window: {
            localStorage: { getItem: () => null, setItem() {} },
            sessionStorage: { getItem: () => null, setItem() {} },
            location: { reload() {} },
            setTimeout: (fn) => fn(),
        },
    };
    vm.createContext(sandbox);
    vm.runInContext(source, sandbox, { filename: "frontend_core.js" });
    sandbox.outputElement = output;
    return sandbox;
}

function citation(key, authors, year) {
    return {
        citationKey: key,
        item: {
            creators: authors.map((lastName) => ({ lastName: lastName })),
            date: year,
        },
    };
}

async function main() {
    /* 1. inserting records the citation on the shape and the slide, in plain text */
    console.log("insert");
    let world = deck();
    let pane = loadPane(world);
    await pane.insertCitationsIntoPowerPoint([citation("zhu2022", ["Zhu"], "2022")]);
    check("the citation is written as plain text", world.text() === " (Zhu, 2022)", JSON.stringify(world.text()));
    check(
        "the slide tag lists the key as recorded",
        JSON.stringify(world.slideTag()) === JSON.stringify([{ k: "zhu2022", l: "Zhu, 2022", r: 1 }]),
        JSON.stringify(world.slideTag()),
    );
    check(
        "the shape records the citation",
        JSON.stringify(world.shapeTag(world.selected)) === JSON.stringify([{ k: "zhu2022", l: "Zhu, 2022" }]),
        JSON.stringify(world.shapeTag(world.selected)),
    );

    /* 2. editing the wording keeps the key */
    console.log("editing the citation");
    world.setText(" (Smyth et al., 2022)");
    await pane.displayCitationsFromSlide();
    check("an edited author keeps the key", world.slideTag().length === 1, JSON.stringify(world.slideTag()));
    world.setText(" (Zhu, 1999)");
    await pane.displayCitationsFromSlide();
    check("an edited year keeps the key", world.slideTag().length === 1, JSON.stringify(world.slideTag()));
    world.setText(" (Zhu, 2022, p. 42)");
    await pane.displayCitationsFromSlide();
    check("an added page number keeps the key", world.slideTag().length === 1, JSON.stringify(world.slideTag()));

    /* 3. deleting the citation drops the key */
    console.log("deleting the citation");
    world.setText("");
    await pane.displayCitationsFromSlide();
    check("an emptied text box drops the key", world.slideTag().length === 0, JSON.stringify(world.slideTag()));
    check("the dropped key is named in the pane", pane.outputElement.children.length > 0);

    /* 4. removing with x takes the text with the key, after an edit */
    console.log("removing with x");
    world = deck();
    pane = loadPane(world);
    await pane.insertCitationsIntoPowerPoint([citation("zhu2022", ["Zhu"], "2022")]);
    world.setText(" (Smyth et al., 2022) done");
    await pane.removeCitation("zhu2022");
    check("the edited citation's text goes with the key", world.text() === "done", JSON.stringify(world.text()));
    check("the slide tag no longer lists the key", world.slideTag().length === 0, JSON.stringify(world.slideTag()));

    /* 5. removing a citation that was already deleted leaves the key alone (nothing half done) */
    console.log("removing a citation whose text is gone");
    world = deck();
    pane = loadPane(world);
    await pane.insertCitationsIntoPowerPoint([citation("zhu2022", ["Zhu"], "2022")]);
    world.setText("nothing left here");
    await pane.removeCitation("zhu2022");
    check("a citation whose text is already gone loses its key", world.slideTag().length === 0, JSON.stringify(world.slideTag()));
    check("the pane says what happened", pane.outputElement.children.length > 0);

    /* 6. citations from an older version are never judged by their text */
    console.log("citations from an older version");
    world = deck();
    pane = loadPane(world);
    world.slide.tags.add("ZOTERO_CITATION_KEYS", JSON.stringify([{ k: "old2019", l: "Old, 2019" }]));
    world.setText("some text without the citation");
    await pane.displayCitationsFromSlide();
    check("an unrecorded citation is left alone", world.slideTag().length === 1, JSON.stringify(world.slideTag()));

    /* 7. two citations in one text box: removing one leaves the other */
    console.log("two citations in one text box");
    world = deck();
    pane = loadPane(world);
    await pane.insertCitationsIntoPowerPoint([citation("zhu2022", ["Zhu"], "2022"), citation("li2021", ["Li"], "2021")]);
    check("both are written", world.text() === " (Zhu, 2022; Li, 2021)", JSON.stringify(world.text()));
    await pane.removeCitation("zhu2022");
    check("only the removed citation's text goes", world.text() === "(Li, 2021)", JSON.stringify(world.text()));
    check(
        "only the removed key goes",
        JSON.stringify(world.slideTag().map((entry) => entry.k)) === JSON.stringify(["li2021"]),
        JSON.stringify(world.slideTag()),
    );

    /* 8. deleting the text box that held the citation drops its key */
    console.log("deleting the text box");
    world = deck();
    pane = loadPane(world);
    await pane.insertCitationsIntoPowerPoint([citation("zhu2022", ["Zhu"], "2022")]);
    world.slide.shapes.items = [];
    await pane.displayCitationsFromSlide();
    check("a deleted text box drops the citation it held", world.slideTag().length === 0, JSON.stringify(world.slideTag()));

    /* 8b. a citation with something added by hand goes as a whole */
    console.log("a citation with a note added by hand");
    world = deck();
    pane = loadPane(world);
    await pane.insertCitationsIntoPowerPoint([citation("zipori2015", ["Zipori"], "2015")]);
    world.setText(" (Zipori et al., 2015, see page 2)");
    await pane.removeCitation("zipori2015");
    check("the whole citation, note included, is removed", world.text() === "", JSON.stringify(world.text()));

    /* 9. a citation rewritten beyond recognition: the key goes, the text stays, the pane says so */
    console.log("a citation rewritten beyond recognition");
    world = deck();
    pane = loadPane(world);
    await pane.insertCitationsIntoPowerPoint([citation("zhu2022", ["Zhu"], "2022")]);
    await pane.removeCitation("zhu2022");
    world = deck();
    pane = loadPane(world);
    await pane.insertCitationsIntoPowerPoint([citation("zhu2022", ["Zhu"], "2022")]);
    world.setText(" (something else entirely)");
    await pane.removeCitation("zhu2022");
    check("a rewritten citation loses its key", world.slideTag().length === 0, JSON.stringify(world.slideTag()));
    check("its text is left alone", world.text() === " (something else entirely)", JSON.stringify(world.text()));

    console.log("");
    console.log("Passed: " + passed + "   Failed: " + failures.length);
    if (failures.length > 0) {
        console.log("Failures:");
        failures.forEach((failure) => console.log("  - " + failure));
        process.exitCode = 1;
    }
}

main().catch((error) => {
    console.error(error);
    process.exitCode = 1;
});
