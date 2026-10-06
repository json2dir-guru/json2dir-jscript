# json2dir-jscript

[json2dir](https://github.com/kitsunoff/awesome-json2dir) in classic Windows JScript (ES3, Windows Script Host), run with Wine's `cscript`; `["link", …]` and `["script", …]` are rejected (RFC §10).

    wine cscript //nologo json2dir.js < document.json

Under Wine, launch through a wrapper that pins `LC_ALL=C.UTF-8` and exports `WINE_UNIX_PID=$$` before `exec wine`: Wine's WSH has no `WScript.StdIn`/`StdErr`, so the script uses `Z:\proc\<pid>\fd\0` and `fd\2`.

Wine's FileSystemObject cannot detect symbolic links, so any entry that already exists (file, folder or link) is refused with an error before anything is written for it; nothing is ever deleted or merged into.
