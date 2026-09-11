1. add bibliography style
    - [x] pass style to server
    - [x] select style in ui
2. select slidemaster while creating new slide
    - [x] use slide master
    - [ ] do not hardcode slide master id
3. format bibliography correctly (italics, subscript, superscript, etc)
    - bbt json rpc return is string, not formatted
    - perhaps, use an auto-exported bibtex file to generate the bibliography
4. migrate to a standalone office addin
    - [x] plain HTML/JS pane served by the helper server (no npm, no webpack, no build step)
    - [x] one-command install: install.ps1 (certificate + sideload registration + autostart)
5. get rid of external python server, perhaps a node.js server?
    - the helper server is standard-library-only Python today; a compiled binary would remove Python as a prerequisite
- [ ] Better icons/author credits
- [ ] Publishing to Microsoft AppSource. See [instructions](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish) for details.
