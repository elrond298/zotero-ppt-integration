1. add bibliography style
    - [x] pass style to server
    - [x] select style in ui
2. select slidemaster while creating new slide
    - [x] use slide master (searches all masters for a "title and content" layout, falls back to the first layout)
    - [x] no hardcoded slide master id
3. format bibliography correctly (italics, subscript, superscript, etc)
    - [x] ask BBT for contentType html and map i/b/em/strong/sub/sup/small-caps onto real PowerPoint runs
4. migrate to a standalone office addin
    - [x] plain HTML/JS pane served by the helper (no npm, no webpack, no build step)
    - [x] one-command install: install.ps1 (certificate + sideload registration + autostart)
5. get rid of the extra runtime dependency
    - [x] Python replaced by ZoteroHelper.cs, compiled with the C# compiler built into Windows
    - [x] no Python, no Node.js, no downloads at install or run time
6. bibliography workflow
    - [x] refresh the References slide in place instead of adding one per click
    - [x] feed the Zotero picker's prefix/locator/suffix/suppress-author fields into the citation text
    - [x] report citation keys Zotero cannot resolve (renamed or deleted items)
    - [x] one line per entry (the whitespace between BBT's block tags no longer becomes a blank line)
    - [x] spread a long bibliography over `References (cont.)` slides, measured with PowerPoint's own autosize,
    - [ ] offer "follow Zotero's quick-copy style" plus a free-text CSL style id
    - [ ] copy BibTeX / CSL-JSON / notes of the cited items (BBT item.export, item.notes)
    - [ ] show author-year next to each citekey in the pane and click through to the citing slides
- [ ] Better icons/author credits
- [ ] Publishing to Microsoft AppSource. See [instructions](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish) for details.
    - needs public HTTPS hosting of the pane, and the pane could then no longer reach the local helper
