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
    - [x] plain HTML/JS pane served by the helper (no npm, no webpack, no build step)
    - [x] one-command install: install.ps1 (certificate + sideload registration + autostart)
5. get rid of the extra runtime dependency
    - [x] Python replaced by ZoteroHelper.cs, compiled with the C# compiler built into Windows
    - [x] no Python, no Node.js, no downloads at install or run time
- [ ] Better icons/author credits
- [ ] Publishing to Microsoft AppSource. See [instructions](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish) for details.
    - needs public HTTPS hosting of the pane, and the pane could then no longer reach the local helper
