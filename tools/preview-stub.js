/* Fake PowerPoint only for rendering documentation screenshots. */
(function () {
    const tagList = (entries) => {
        const items = (entries || []).map(([key, value]) => ({ key, value }));
        return { items, load() {}, add(key, value) { const f = items.find((t) => t.key === key); if (f) f.value = value; else items.push({ key, value }); } };
    };
    const shape = (id, text) => ({ id: "shape-" + id, name: "Content Placeholder", textFrame: { textRange: { text, load() {}, font: {} } }, tags: tagList() });

    const slide2 = { tags: tagList([["ZOTERO_CITATION_KEYS", JSON.stringify([
        { k: "zuberiHeterogeneousFreezingAqueous2001", l: "Zuberi et al., 2001", s: "shape-2" },
        { k: "ziporiEffectsAerosolSources2015", l: "Zipori et al., 2015", s: "shape-2" }])]]),
        shapes: { items: [shape(2, "Ice nucleation measurements (Zuberi et al., 2001; Zipori et al., 2015)")], load() {} } };
    const slide1 = { tags: tagList([["ZOTERO_CITATION_KEYS", JSON.stringify([
        { k: "zhouSeasonalVariabilitySources2026", l: "Zhou et al., 2026", s: "shape-1" }])]]),
        shapes: { items: [shape(1, "Aerosol sources (Zhou et al., 2026)")], load() {} } };

    const context = {
        presentation: {
            slides: { items: [slide1, slide2], load() {} },
            getSelectedSlides: () => ({ items: [slide2], load() {}, getItemAt: () => slide2 }),
            getSelectedTextRange: () => slide2.shapes.items[0].textFrame.textRange,
        },
        sync: () => Promise.resolve(),
    };

    window.Office = {
        onReady: (callback) => setTimeout(() => callback({ host: "PowerPoint" }), 0),
        context: {
            requirements: { isSetSupported: () => true },
            diagnostics: { host: "PowerPoint", platform: "PC", version: "16.0.17932" },
            document: { addHandlerAsync: () => {} },
        },
        EventType: { DocumentSelectionChanged: "documentSelectionChanged" },
    };
    window.PowerPoint = { run: (callback) => Promise.resolve(callback(context)), ShapeAutoSize: { autoSizeShapeToFitText: 0 } };
    window.fetch = (url) => {
        if (String(url).indexOf("/health") >= 0) return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({ status: "ok" }) });
        return Promise.resolve({ ok: true, status: 200, text: () => Promise.resolve(""), json: () => Promise.resolve({}) });
    };
})();
