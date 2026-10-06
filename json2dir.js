// json2dir: create the directory tree a JSON document describes, in the current directory.
// Implements RFC J2D-1 (https://github.com/kitsunoff/awesome-json2dir/blob/main/spec/rfc-json2dir.md)
// in classic JScript (ES3) for Windows Script Host:  cscript //nologo json2dir.js < document.json
// Under Wine, run it through a launcher that exports WINE_UNIX_PID (see README): Wine's WSH has
// no WScript.StdIn/StdErr, so the script reads its Unix stdin/stderr via Z:\proc\<pid>\fd\N.

var fso = new ActiveXObject("Scripting.FileSystemObject");

function Fail(message) { this.message = message; }
function fail(message) { throw new Fail(message); }

// ---------- standard streams ----------

var unixPid = "";
try { unixPid = new ActiveXObject("WScript.Shell").Environment("PROCESS")("WINE_UNIX_PID"); } catch (e) {}

function writeStderr(text) {
  try {
    if (unixPid) {
      var t = fso.OpenTextFile("Z:\\proc\\" + unixPid + "\\fd\\2", 2, false, 0);
      t.Write(text);
      t.Close();
    } else {
      WScript.StdErr.Write(text);
    }
  } catch (e) {}
}

// Text streams in ANSI mode map each byte through code page 1252 (the launcher pins the locale).
// These tables turn such a string back into bytes and bytes into such a string.
var CP1252_HIGH = [0x20AC, 0x81, 0x201A, 0x192, 0x201E, 0x2026, 0x2020, 0x2021, 0x2C6, 0x2030,
  0x160, 0x2039, 0x152, 0x8D, 0x17D, 0x8F, 0x90, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013,
  0x2014, 0x2DC, 0x2122, 0x161, 0x203A, 0x153, 0x9D, 0x17E, 0x178];
var byteToChar = [], charToByte = {};
for (var b = 0; b < 256; b++) {
  var c = (b >= 0x80 && b < 0xA0) ? CP1252_HIGH[b - 0x80] : b;
  byteToChar[b] = String.fromCharCode(c);
  charToByte[c] = b;
}

// Returns the input as an array of byte values.
function readStdinBytes() {
  var s = "";
  if (unixPid) {
    var t = fso.OpenTextFile("Z:\\proc\\" + unixPid + "\\fd\\0", 1, false, 0);
    if (!t.AtEndOfStream) s = t.ReadAll();
    t.Close();
  } else {
    if (!WScript.StdIn.AtEndOfStream) s = WScript.StdIn.ReadAll();
  }
  var bytes = [];
  for (var i = 0; i < s.length; i++) {
    var v = charToByte[s.charCodeAt(i)];
    if (v === undefined) fail("cannot read standard input as bytes");
    bytes[i] = v;
  }
  return bytes;
}

// ---------- UTF-8 ----------

// §3: strict UTF-8 -> JScript (UTF-16) string.
function decodeUtf8(bytes) {
  var out = [], i = 0, n = bytes.length;
  if (n >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) i = 3; // BOM may be ignored
  while (i < n) {
    var b0 = bytes[i], cp, need, min;
    if (b0 < 0x80) { out.push(String.fromCharCode(b0)); i++; continue; }
    if (b0 >= 0xC2 && b0 <= 0xDF) { cp = b0 & 0x1F; need = 1; min = 0x80; }
    else if (b0 >= 0xE0 && b0 <= 0xEF) { cp = b0 & 0x0F; need = 2; min = 0x800; }
    else if (b0 >= 0xF0 && b0 <= 0xF4) { cp = b0 & 0x07; need = 3; min = 0x10000; }
    else fail("input is not valid UTF-8");
    for (var k = 1; k <= need; k++) {
      var bk = bytes[i + k];
      if (bk === undefined || (bk & 0xC0) != 0x80) fail("input is not valid UTF-8");
      cp = cp * 64 + (bk & 0x3F);
    }
    if (cp < min || cp > 0x10FFFF || (cp >= 0xD800 && cp <= 0xDFFF)) fail("input is not valid UTF-8");
    if (cp >= 0x10000) {
      cp -= 0x10000;
      out.push(String.fromCharCode(0xD800 + (cp >> 10), 0xDC00 + (cp & 0x3FF)));
    } else out.push(String.fromCharCode(cp));
    i += need + 1;
  }
  return out.join("");
}

// JScript string (already checked for lone surrogates) -> string of byte chars 0..255.
function utf8Bytes(s) {
  var out = [];
  for (var i = 0; i < s.length; i++) {
    var c = s.charCodeAt(i);
    if (c >= 0xD800 && c <= 0xDBFF) { c = 0x10000 + (c - 0xD800) * 1024 + (s.charCodeAt(++i) - 0xDC00); }
    if (c < 0x80) out.push(c);
    else if (c < 0x800) out.push(0xC0 | (c >> 6), 0x80 | (c & 63));
    else if (c < 0x10000) out.push(0xE0 | (c >> 12), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
    else out.push(0xF0 | (c >> 18), 0x80 | ((c >> 12) & 63), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
  }
  return out;
}

// ---------- JSON (RFC 8259, strict) ----------
// Objects become {names: [...], values: {"$"+name: value}}, arrays become JScript arrays,
// strings stay strings; numbers/true/false/null become {bad: "number"} etc. so §4.5 can report them.

function parseJson(text) {
  var pos = 0;
  function err(what) { fail("input is not valid JSON: " + what + " at offset " + pos); }
  function ws() {
    for (;;) {
      var c = text.charAt(pos);
      if (c == " " || c == "\t" || c == "\n" || c == "\r") pos++; else return;
    }
  }
  function value() {
    ws();
    var c = text.charAt(pos);
    if (c == "{") return object();
    if (c == "[") return array();
    if (c == "\"") return string();
    if (c == "t") return literal("true");
    if (c == "f") return literal("false");
    if (c == "n") return literal("null");
    if (c == "-" || (c >= "0" && c <= "9")) return number();
    err(pos >= text.length ? "unexpected end" : "unexpected character");
  }
  function literal(word) {
    if (text.substr(pos, word.length) != word) err("bad literal");
    pos += word.length;
    return { bad: word };
  }
  function number() {
    var m = /^-?(0|[1-9][0-9]*)(\.[0-9]+)?([eE][+-]?[0-9]+)?/.exec(text.substr(pos, 400));
    if (!m) err("bad number");
    pos += m[0].length;
    return { bad: "number" };
  }
  function string() {
    pos++; // opening quote
    var out = [], start = pos;
    for (;;) {
      if (pos >= text.length) err("unterminated string");
      var code = text.charCodeAt(pos);
      if (code == 34) { out.push(text.substring(start, pos)); pos++; return out.join(""); }
      if (code < 0x20) err("control character in string");
      if (code == 92) {
        out.push(text.substring(start, pos));
        var e = text.charAt(pos + 1);
        pos += 2;
        if (e == "\"") out.push("\"");
        else if (e == "\\") out.push("\\");
        else if (e == "/") out.push("/");
        else if (e == "b") out.push("\b");
        else if (e == "f") out.push("\f");
        else if (e == "n") out.push("\n");
        else if (e == "r") out.push("\r");
        else if (e == "t") out.push("\t");
        else if (e == "u") {
          var h = text.substr(pos, 4);
          if (!/^[0-9a-fA-F]{4}$/.test(h)) err("bad \\u escape");
          out.push(String.fromCharCode(parseInt(h, 16)));
          pos += 4;
        } else err("bad escape");
        start = pos;
      } else pos++;
    }
  }
  function array() {
    pos++;
    var arr = [];
    ws();
    if (text.charAt(pos) == "]") { pos++; return arr; }
    for (;;) {
      arr.push(value());
      ws();
      var c = text.charAt(pos++);
      if (c == "]") return arr;
      if (c != ",") { pos--; err("expected , or ]"); }
    }
  }
  function object() {
    pos++;
    var obj = { names: [], values: {} };
    ws();
    if (text.charAt(pos) == "}") { pos++; return obj; }
    for (;;) {
      ws();
      if (text.charAt(pos) != "\"") err("expected a member name");
      var name = string();
      ws();
      if (text.charAt(pos) != ":") err("expected :");
      pos++;
      var v = value();
      // Duplicate names: the last occurrence wins (§4.2).
      if (!obj.values.hasOwnProperty("$" + name)) obj.names.push(name);
      obj.values["$" + name] = v;
      ws();
      var c = text.charAt(pos++);
      if (c == "}") return obj;
      if (c != ",") { pos--; err("expected , or }"); }
    }
  }
  var result = value();
  ws();
  if (pos < text.length) err("trailing data");
  return result;
}

function isObject(v) { return v !== null && typeof v == "object" && v.names !== undefined; }
function isArray(v) { return v instanceof Array; }

// ---------- validation (§4, before touching the file system) ----------

function hasLoneSurrogate(s) {
  for (var i = 0; i < s.length; i++) {
    var c = s.charCodeAt(i);
    if (c >= 0xD800 && c <= 0xDBFF) {
      var d = s.charCodeAt(i + 1);
      if (d >= 0xDC00 && d <= 0xDFFF) { i++; continue; }
      return true;
    }
    if (c >= 0xDC00 && c <= 0xDFFF) return true;
  }
  return false;
}

function checkString(s, where) {
  if (hasLoneSurrogate(s)) fail(where + ": string contains an unpaired surrogate");
}

function show(s) { return "\"" + s + "\""; }

function validate(v, where) {
  if (typeof v == "string") return checkString(v, where);
  if (isArray(v)) {
    if (v.length != 2 || typeof v[0] != "string" || typeof v[1] != "string")
      fail(where + ": an array must be [\"link\", target] or [\"script\", content]");
    if (v[0] != "link" && v[0] != "script") fail(where + ": unknown array kind " + show(v[0]));
    // §10: WSH can neither create symbolic links nor set execute bits, so reject, never approximate.
    if (v[0] == "link") fail(where + ": symbolic links cannot be created from JScript/WSH (RFC section 10)");
    fail(where + ": executable files cannot be created from JScript/WSH (RFC section 10)");
  }
  if (isObject(v)) {
    var seen = {};
    for (var i = 0; i < v.names.length; i++) {
      var name = v.names[i];
      var child = where == "." ? name : where + "/" + name;
      checkString(name, child);
      if (name == "" || name == "." || name == ".." || name.indexOf("/") >= 0 || name.indexOf("\u0000") >= 0)
        fail(child + ": invalid name " + show(name));
      // Win32 paths treat these as separators or wildcards; they cannot be used safely in a name.
      if (/[\\*?"<>|]/.test(name)) fail(child + ": the name cannot be represented through Win32 file APIs");
      // Wine's file lookup is case-insensitive: names differing only in case would collide.
      var folded = "$" + name.toLowerCase();
      if (seen.hasOwnProperty(folded)) fail(child + ": names differing only in case are not supported here");
      seen[folded] = true;
      validate(v.values["$" + name], child);
    }
    return;
  }
  fail(where + ": " + v.bad + " values are not allowed");
}

// ---------- creating the tree (§4, §5) ----------

function byteString(s) {
  var b = utf8Bytes(s), out = [];
  for (var i = 0; i < b.length; i++) out.push(String.fromCharCode(b[i]));
  return out.join("");
}

function compareUtf8(a, b) {
  var x = byteString(a), y = byteString(b);
  return x < y ? -1 : x > y ? 1 : 0;
}

function writeFile(path, content, where) {
  var b = utf8Bytes(content), chars = [];
  for (var i = 0; i < b.length; i++) chars.push(byteToChar[b[i]]);
  var t = fso.CreateTextFile(path, false, false); // fails if the name exists (e.g. a dangling symlink)
  if (chars.length) t.Write(chars.join(""));
  t.Close();
}

function apply(dir, rel, tree) {
  var names = tree.names.slice(0).sort(compareUtf8);
  for (var i = 0; i < names.length; i++) {
    var name = names[i], value = tree.values["$" + name];
    var path = dir + "\\" + name, where = rel == "" ? name : rel + "/" + name;
    try {
      var isDir = fso.FolderExists(path), isFile = fso.FileExists(path);
      if (typeof value == "string") {
        if (isDir) fail(where + ": a directory is in the way");
        if (isFile) fso.DeleteFile(path, true);
        writeFile(path, value, where);
      } else {
        if (!isDir) {
          if (isFile) fso.DeleteFile(path, true);
          fso.CreateFolder(path);
        }
        apply(path, where, value);
      }
    } catch (e) {
      if (e instanceof Fail) throw e;
      fail(where + ": " + (e.message || "file system error") + " (0x" + (e.number >>> 0).toString(16) + ")");
    }
  }
}

function main() {
  if (WScript.Arguments.length > 0) {
    writeStderr("usage: json2dir < document.json\n");
    return 2;
  }
  try {
    var doc = parseJson(decodeUtf8(readStdinBytes()));
    if (!isObject(doc)) fail("the root of the document must be an object");
    validate(doc, ".");
    // "\\?\" turns off Win32 path rewriting, so names like "a." or "con" are used as given.
    apply("\\\\?\\" + fso.GetAbsolutePathName("."), "", doc);
    return 0;
  } catch (e) {
    writeStderr("json2dir: " + (e instanceof Fail ? e.message : (e.message || "error") + " (0x" + (e.number >>> 0).toString(16) + ")") + "\n");
    return 1;
  }
}

WScript.Quit(main());
