/**
 * WordPasteEditor – minimal contenteditable HTML editor with Word-paste support.
 * Zero dependencies (uses WordCleaner from word-cleaner.js).
 *
 * @param {string|HTMLElement} element - The contenteditable element or CSS selector.
 * @param {Object} [options]
 * @param {string|HTMLElement} [options.sourceView] - Textarea element or selector for HTML source view.
 * @param {string|HTMLElement} [options.sourceToggle] - Button element or selector to toggle source view.
 * @param {string} [options.placeholder] - Placeholder text for the editor.
 * @param {boolean|string[]} [options.toolbar] - Show a formatting toolbar above the
 *   editor: true for every button, or an array of button names ("|" for a separator).
 *   Needs toolbar.css. Off by default.
 * @param {string[]} [options.linkSchemes] - Link schemes the toolbar accepts.
 *   Default ["http", "https"].
 */
(function (root, factory) {
  if (typeof define === "function" && define.amd) {
    define(["./word-cleaner"], factory);
  } else if (typeof module === "object" && module.exports) {
    module.exports = factory(require("./word-cleaner"));
  } else {
    root.WordPasteEditor = factory(root.WordCleaner);
  }
})(typeof self !== "undefined" ? self : this, function (WordCleaner) {
  "use strict";

  function WordPasteEditor(element, options) {
    if (!(this instanceof WordPasteEditor)) {
      return new WordPasteEditor(element, options);
    }

    this._editor = resolve(element);
    if (!this._editor) throw new Error("WordPasteEditor: element not found");

    options = options || {};
    this._isSourceMode = false;

    if (!this._editor.isContentEditable) {
      this._editor.contentEditable = "true";
    }

    if (options.placeholder && !this._editor.getAttribute("data-placeholder")) {
      this._editor.setAttribute("data-placeholder", options.placeholder);
    }

    this._sourceView = options.sourceView
      ? resolve(options.sourceView)
      : null;
    this._toggle = options.sourceToggle
      ? resolve(options.sourceToggle)
      : null;

    this._handlers = {};
    this._setupPaste();
    this._setupDrop();
    this._setupKeyboard();

    this._toolbar = null;
    if (options.toolbar) {
      this._toolbar = createToolbar(
        this._editor,
        options.toolbar === true ? DEFAULT_TOOLBAR : options.toolbar,
        (options.linkSchemes || ["http", "https"]).map(function (s) {
          return String(s).toLowerCase();
        })
      );
    }

    if (this._toggle && this._sourceView) {
      this._setupSourceToggle();
    }

    this._editor.focus();
  }

  function resolve(el) {
    return typeof el === "string" ? document.querySelector(el) : el;
  }

  // ────────── Paste handling ─────────────────────────────────

  WordPasteEditor.prototype._setupPaste = function () {
    var self = this;
    this._handlers.paste = function (e) {
      var clipboardData = e.clipboardData || window.clipboardData;
      if (!clipboardData) return;

      var html = clipboardData.getData("text/html");
      var plainText = clipboardData.getData("text/plain");

      if (html) {
        e.preventDefault();
        self._insertHTML(WordCleaner.clean(html));
        return;
      }

      if (plainText && !html) {
        e.preventDefault();
        var escaped = escapeHTML(plainText).replace(/\n/g, "<br>");
        self._insertHTML(escaped);
      }
    };
    this._editor.addEventListener("paste", this._handlers.paste);
  };

  // ────────── Insertion helper ───────────────────────────────

  WordPasteEditor.prototype._insertHTML = function (html) {
    if (
      document.queryCommandSupported &&
      document.queryCommandSupported("insertHTML")
    ) {
      document.execCommand("insertHTML", false, html);
      return;
    }

    var sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) {
      this._editor.innerHTML += html;
      return;
    }

    var range = sel.getRangeAt(0);
    range.deleteContents();

    var frag = document.createRange().createContextualFragment(html);
    var lastNode = frag.lastChild;
    range.insertNode(frag);

    if (lastNode) {
      var newRange = document.createRange();
      newRange.setStartAfter(lastNode);
      newRange.collapse(true);
      sel.removeAllRanges();
      sel.addRange(newRange);
    }
  };

  function escapeHTML(text) {
    var div = document.createElement("div");
    div.appendChild(document.createTextNode(text));
    return div.innerHTML;
  }

  // ────────── Source view toggle ─────────────────────────────

  WordPasteEditor.prototype._setupSourceToggle = function () {
    var self = this;
    this._handlers.toggle = function () {
      self._isSourceMode = !self._isSourceMode;

      if (self._isSourceMode) {
        self._sourceView.value = formatHTML(self._editor.innerHTML);
        self._editor.style.display = "none";
        self._sourceView.style.display = "block";
        self._toggle.textContent = "View Editor";
        self._sourceView.focus();
      } else {
        self._editor.innerHTML = self._sourceView.value;
        self._sourceView.style.display = "none";
        self._editor.style.display = "block";
        self._toggle.textContent = "View Source";
        self._editor.focus();
      }

      if (self._toolbar) setToolbarDisabled(self._toolbar, self._isSourceMode);
    };
    this._toggle.addEventListener("click", this._handlers.toggle);
  };

  // ────────── Simple HTML formatter for source view ──────────

  /**
   * Very basic HTML indentation for readability in the source view.
   * Not a full pretty-printer, but good enough to read.
   */
  function formatHTML(html) {
    var result = "";
    var indent = 0;
    var indentStr = "  ";

    // Normalize
    html = html.replace(/>\s+</g, "><").trim();

    // Split on tags
    var tokens = html.split(/(<\/?[^>]+>)/g);

    for (var i = 0; i < tokens.length; i++) {
      var token = tokens[i].trim();
      if (!token) continue;

      if (token.match(/^<\//)) {
        // Closing tag
        indent = Math.max(0, indent - 1);
        result += repeat(indentStr, indent) + token + "\n";
      } else if (token.match(/^<[^/][^>]*[^/]>$/) && !token.match(/^<(br|hr|img|input|meta|link)\b/i)) {
        // Opening tag (not self-closing, not void)
        result += repeat(indentStr, indent) + token + "\n";
        indent++;
      } else if (token.match(/^</)) {
        // Self-closing or void tag
        result += repeat(indentStr, indent) + token + "\n";
      } else {
        // Text content
        result += repeat(indentStr, indent) + token + "\n";
      }
    }

    return result.trim();
  }

  function repeat(str, times) {
    var result = "";
    for (var i = 0; i < times; i++) result += str;
    return result;
  }

  // ────────── Drop handling ──────────────────────────────────

  WordPasteEditor.prototype._setupDrop = function () {
    var self = this;
    this._handlers.drop = function (e) {
      var dt = e.dataTransfer;
      if (!dt) return;

      var html = dt.getData("text/html");
      if (html) {
        e.preventDefault();
        var cleaned = WordCleaner.clean(html);
        var range;
        if (document.caretRangeFromPoint) {
          range = document.caretRangeFromPoint(e.clientX, e.clientY);
        } else if (e.rangeParent) {
          range = document.createRange();
          range.setStart(e.rangeParent, e.rangeOffset);
        }

        if (range) {
          var sel = window.getSelection();
          sel.removeAllRanges();
          sel.addRange(range);
        }

        self._insertHTML(cleaned);
      }
    };

    this._handlers.dragover = function (e) {
      e.preventDefault();
    };

    this._editor.addEventListener("drop", this._handlers.drop);
    this._editor.addEventListener("dragover", this._handlers.dragover);
  };

  // ────────── Keyboard shortcuts ─────────────────────────────

  WordPasteEditor.prototype._setupKeyboard = function () {
    var self = this;
    this._handlers.keydown = function (e) {
      if (e.key === "Tab") {
        e.preventDefault();
        self._insertHTML("&nbsp;&nbsp;&nbsp;&nbsp;");
      }
    };
    this._editor.addEventListener("keydown", this._handlers.keydown);
  };

  // ────────── Formatting toolbar (opt-in) ────────────────────

  /**
   * Toolbar buttons by name. "cmd" buttons run document.execCommand, and
   * "state" buttons show a pressed state that follows the selection.
   */
  var BUTTONS = {
    bold: { cmd: "bold", title: "Bold (Ctrl+B)", state: true },
    italic: { cmd: "italic", title: "Italic (Ctrl+I)", state: true },
    underline: { cmd: "underline", title: "Underline (Ctrl+U)", state: true },
    color: { color: "foreColor", title: "Text colour", initial: "#c00000" },
    highlight: { color: "hiliteColor", title: "Highlight colour", initial: "#ffff00", none: true },
    alignLeft: { cmd: "justifyLeft", title: "Align left", state: true },
    alignCenter: { cmd: "justifyCenter", title: "Centre", state: true },
    alignRight: { cmd: "justifyRight", title: "Align right", state: true },
    justify: { cmd: "justifyFull", title: "Justify", state: true },
    link: { title: "Insert link" },
    unlink: { title: "Remove link" },
    clear: { cmd: "removeFormat", title: "Clear formatting" }
  };

  /** Every button, in order. "|" is a separator. */
  var DEFAULT_TOOLBAR = [
    "bold", "italic", "underline", "|",
    "color", "highlight", "|",
    "alignLeft", "alignCenter", "alignRight", "justify", "|",
    "link", "unlink", "|",
    "clear"
  ];

  var PALETTE = [
    "#000000", "#262626", "#404040", "#595959", "#7f7f7f", "#a6a6a6", "#bfbfbf", "#d9d9d9", "#f2f2f2", "#ffffff",
    "#c00000", "#ff0000", "#ffc000", "#ffff00", "#92d050", "#00b050", "#00b0f0", "#0070c0", "#002060", "#7030a0",
    "#f2dcdb", "#fde9d9", "#fff2cc", "#ffffcc", "#ebf1de", "#e2efda", "#daeef3", "#dce6f1", "#d9e1f2", "#e4dfec"
  ];

  function svg(paths) {
    return '<svg viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path d="' + paths + '"/></svg>';
  }

  var ICONS = {
    bold: '<span class="wpe-tb-glyph" style="font-weight:700">B</span>',
    italic: '<span class="wpe-tb-glyph" style="font-style:italic">I</span>',
    underline: '<span class="wpe-tb-glyph" style="text-decoration:underline">U</span>',
    color: '<span class="wpe-tb-glyph">A</span><span class="wpe-tb-bar"></span>',
    highlight: svg("M10 2.5l3.5 3.5-6 6H4v-3.5zM8.5 4l3.5 3.5") + '<span class="wpe-tb-bar"></span>',
    alignLeft: svg("M2 3h12M2 6.5h8M2 10h12M2 13.5h8"),
    alignCenter: svg("M2 3h12M4 6.5h8M2 10h12M4 13.5h8"),
    alignRight: svg("M2 3h12M6 6.5h8M2 10h12M6 13.5h8"),
    justify: svg("M2 3h12M2 6.5h12M2 10h12M2 13.5h12"),
    link: svg("M6.5 9.5l3-3M7 4.5l1-1a2.8 2.8 0 0 1 4 4l-1 1M9 11.5l-1 1a2.8 2.8 0 0 1-4-4l1-1"),
    unlink: svg("M7 4.5l1-1a2.8 2.8 0 0 1 4 4l-1 1M9 11.5l-1 1a2.8 2.8 0 0 1-4-4l1-1M2.5 2.5l2 2M11.5 11.5l2 2"),
    clear: svg("M3 3.5h8M7 3.5v9M10 9.5l3.5 3.5M13.5 9.5L10 13")
  };

  /**
   * Build the toolbar, insert it just before the editor element and start
   * listening. Formatting goes through document.execCommand, which fires
   * the editor's "input" event like typing does.
   */
  function createToolbar(el, names, schemes) {
    var st = { el: el, schemes: schemes, range: null, open: null, buttons: {}, bars: {}, handlers: {} };

    st.bar = document.createElement("div");
    st.bar.className = "wpe-toolbar";
    st.bar.setAttribute("role", "toolbar");
    st.bar.setAttribute("aria-label", "Text formatting");
    if (el.style.width) st.bar.style.width = el.style.width;

    for (var i = 0; i < names.length; i++) {
      if (names[i] === "|") addSeparator(st.bar);
      else if (BUTTONS.hasOwnProperty(names[i])) addToolbarButton(st, names[i]);
    }
    el.parentNode.insertBefore(st.bar, el);

    // Remember the last selection inside the editor, so a toolbar action
    // still applies to it after focus moved into a popover.
    st.handlers.selectionchange = function () {
      var sel = window.getSelection();
      if (!sel || sel.rangeCount === 0 || !el.contains(sel.anchorNode)) return;
      st.range = sel.getRangeAt(0).cloneRange();
      refresh(st);
    };
    // Clicking anywhere outside an open popover closes it.
    st.handlers.mousedown = function (e) {
      if (st.open && !st.open.pop.contains(e.target) && !st.open.button.contains(e.target)) {
        closePopover(st);
      }
    };
    // Ctrl+B and the like change the pressed state without moving the selection.
    st.handlers.keyup = function () { refresh(st); };

    document.addEventListener("selectionchange", st.handlers.selectionchange);
    document.addEventListener("mousedown", st.handlers.mousedown);
    el.addEventListener("keyup", st.handlers.keyup);
    return st;
  }

  function destroyToolbar(st) {
    closePopover(st);
    document.removeEventListener("selectionchange", st.handlers.selectionchange);
    document.removeEventListener("mousedown", st.handlers.mousedown);
    st.el.removeEventListener("keyup", st.handlers.keyup);
    if (st.bar.parentNode) st.bar.parentNode.removeChild(st.bar);
  }

  /** Grey the buttons out while the source view is showing. */
  function setToolbarDisabled(st, disabled) {
    closePopover(st);
    for (var name in st.buttons) st.buttons[name].disabled = disabled;
  }

  function addToolbarButton(st, name) {
    var def = BUTTONS[name];
    var b;
    if (def.color) {
      b = addButton(st, name, function (btn) {
        togglePopover(st, btn, function () { return buildColorPopover(st, name); });
      });
      b.setAttribute("aria-haspopup", "true");
      b.setAttribute("aria-expanded", "false");
      st.bars[name] = b.querySelector(".wpe-tb-bar");
      st.bars[name].style.background = def.initial;
    } else if (name === "link") {
      b = addButton(st, name, function (btn) {
        togglePopover(st, btn, function () { return buildLinkPopover(st); });
      });
      b.setAttribute("aria-haspopup", "true");
      b.setAttribute("aria-expanded", "false");
    } else if (name === "unlink") {
      addButton(st, name, function () { removeLink(st); });
    } else {
      b = addButton(st, name, function () {
        closePopover(st);
        exec(st, def.cmd, null, false);
      });
      if (def.state) b.setAttribute("aria-pressed", "false");
    }
  }

  function addButton(st, name, onClick) {
    var title = BUTTONS[name].title;
    var b = document.createElement("button");
    b.type = "button"; // never submit a surrounding form
    b.className = "wpe-tb-btn wpe-tb-" + name;
    b.title = title;
    b.setAttribute("aria-label", title);
    b.innerHTML = ICONS[name];
    // Keep focus and the selection in the editor while clicking.
    b.addEventListener("mousedown", function (e) { e.preventDefault(); });
    b.addEventListener("click", function (e) { e.preventDefault(); onClick(b); });
    st.bar.appendChild(b);
    st.buttons[name] = b;
    return b;
  }

  function addSeparator(bar) {
    var s = document.createElement("span");
    s.className = "wpe-tb-sep";
    s.setAttribute("aria-hidden", "true");
    bar.appendChild(s);
  }

  // ────────── Toolbar: colours ───────────────────────────────

  function buildColorPopover(st, name) {
    var pop = document.createElement("div");
    pop.className = "wpe-tb-pop";

    var grid = document.createElement("div");
    grid.className = "wpe-tb-swatches";
    PALETTE.forEach(function (c) {
      var s = document.createElement("button");
      s.type = "button";
      s.className = "wpe-tb-swatch";
      s.style.background = c;
      s.title = c;
      s.setAttribute("aria-label", c);
      s.addEventListener("mousedown", function (e) { e.preventDefault(); });
      s.addEventListener("click", function (e) { e.preventDefault(); applyColor(st, name, c); });
      grid.appendChild(s);
    });
    pop.appendChild(grid);

    var row = document.createElement("div");
    row.className = "wpe-tb-pop-row";

    if (BUTTONS[name].none) {
      var none = document.createElement("button");
      none.type = "button";
      none.className = "wpe-tb-textbtn";
      none.textContent = "No colour";
      none.addEventListener("mousedown", function (e) { e.preventDefault(); });
      none.addEventListener("click", function (e) { e.preventDefault(); applyColor(st, name, "transparent"); });
      row.appendChild(none);
    }

    // Native colour picker for anything outside the palette. It takes focus
    // away from the editor, so applyColor restores the saved selection.
    var more = document.createElement("label");
    more.className = "wpe-tb-more";
    var input = document.createElement("input");
    input.type = "color"; // no name attribute: never posted with a form
    input.value = toHex(st.bars[name].style.background) || "#000000";
    input.addEventListener("change", function () { applyColor(st, name, input.value); });
    more.appendChild(input);
    more.appendChild(document.createTextNode(" More colours…"));
    row.appendChild(more);

    pop.appendChild(row);
    return pop;
  }

  function applyColor(st, name, color) {
    closePopover(st);
    var cmd = BUTTONS[name].color;
    if (cmd === "hiliteColor" && !supported("hiliteColor")) cmd = "backColor";
    // styleWithCSS gives <span style="color: ..."> instead of <font>
    exec(st, cmd, color, true);
    if (color !== "transparent") st.bars[name].style.background = color;
  }

  // ────────── Toolbar: links ─────────────────────────────────

  function buildLinkPopover(st) {
    var existing = st.range ? closestLink(st.range.commonAncestorContainer, st.el) : null;

    var pop = document.createElement("div");
    pop.className = "wpe-tb-pop wpe-tb-linkpop";

    var label = document.createElement("label");
    label.className = "wpe-tb-label";
    label.textContent = "Web address";
    var input = document.createElement("input");
    // Not type=url: inside a form, browser validation would block its submit.
    input.type = "text";
    input.setAttribute("inputmode", "url");
    input.setAttribute("autocomplete", "off");
    input.setAttribute("spellcheck", "false");
    input.placeholder = "https://";
    input.value = existing ? existing.getAttribute("href") || "" : "";
    label.appendChild(input);
    pop.appendChild(label);

    var err = document.createElement("div");
    err.className = "wpe-tb-err";
    err.hidden = true;
    pop.appendChild(err);

    var row = document.createElement("div");
    row.className = "wpe-tb-pop-row";
    var ok = document.createElement("button");
    ok.type = "button";
    ok.className = "wpe-tb-textbtn wpe-tb-primary";
    ok.textContent = existing ? "Update link" : "Insert link";
    row.appendChild(ok);
    if (existing) {
      var rm = document.createElement("button");
      rm.type = "button";
      rm.className = "wpe-tb-textbtn";
      rm.textContent = "Remove link";
      rm.addEventListener("click", function (e) { e.preventDefault(); closePopover(st); removeLink(st); });
      row.appendChild(rm);
    }
    pop.appendChild(row);

    function submit() {
      var url = normalizeUrl(input.value, st.schemes);
      if (!url) {
        err.textContent = linkHint(st.schemes);
        err.hidden = false;
        input.focus();
        return;
      }
      closePopover(st);
      insertLink(st, url, existing);
    }
    ok.addEventListener("click", function (e) { e.preventDefault(); submit(); });
    input.addEventListener("keydown", function (e) {
      // Enter would otherwise submit a surrounding form
      if (e.key === "Enter") { e.preventDefault(); submit(); }
    });

    pop._focusTarget = input;
    return pop;
  }

  function insertLink(st, url, existing) {
    restore(st);
    var sel = window.getSelection();
    if (existing && sel.isCollapsed) {
      // A direct DOM edit, so no "input" event unless we send one
      existing.setAttribute("href", url);
      changed(st);
    } else if (sel.isCollapsed) {
      var text = url.replace(/^mailto:/i, "");
      exec(st, "insertHTML", '<a href="' + escapeAttr(url) + '">' + escapeAttr(text) + "</a>", false);
    } else {
      exec(st, "createLink", url, false);
    }
  }

  function removeLink(st) {
    closePopover(st);
    restore(st);
    var sel = window.getSelection();
    if (sel.isCollapsed) {
      var a = closestLink(sel.anchorNode, st.el);
      if (!a) return;
      var r = document.createRange();
      r.selectNodeContents(a);
      sel.removeAllRanges();
      sel.addRange(r);
      st.range = r.cloneRange(); // exec restores st.range before running
    }
    exec(st, "unlink", null, false);
  }

  /**
   * Turn what the user typed into a link, or null when its scheme isn't
   * allowed. A bare domain gets https:// and a bare email address gets
   * mailto:, when those schemes are allowed.
   */
  function normalizeUrl(s, schemes) {
    s = (s || "").trim();
    if (!s || /\s/.test(s)) return null;

    var scheme = /^([a-z][a-z0-9+.\-]*):(?!\d)/i.exec(s); // not host:port
    if (scheme) {
      var name = scheme[1].toLowerCase();
      if (schemes.indexOf(name) < 0) return null; // javascript:, data: ...
      if ((name === "http" || name === "https") && !/^https?:\/\/[^\/?#]+/i.test(s)) return null;
      return s;
    }

    if (/^[^\/?#:@]+@[^\/?#:@]+\.[a-z]{2,}$/i.test(s)) {
      return schemes.indexOf("mailto") >= 0 ? "mailto:" + s : null;
    }

    var web = schemes.indexOf("https") >= 0 ? "https:" : schemes.indexOf("http") >= 0 ? "http:" : null;
    if (!web) return null;
    if (/^\/\//.test(s)) return web + s;
    if (/^[^\/?#:@]+\.[a-z]{2,}([:\/?#]|$)/i.test(s)) return web + "//" + s; // www.example.com/...
    return null;
  }

  /** "Enter an address that starts with http:// or https://" */
  function linkHint(schemes) {
    var parts = schemes.map(function (s) {
      return s === "http" || s === "https" || s === "ftp" ? s + "://" : s + ":";
    });
    var last = parts.pop();
    return "Enter an address that starts with " + (parts.length ? parts.join(", ") + " or " : "") + last;
  }

  function closestLink(node, el) {
    while (node && node !== el) {
      if (node.nodeType === 1 && node.tagName === "A") return node;
      node = node.parentNode;
    }
    return null;
  }

  // ────────── Toolbar: popovers ──────────────────────────────

  function togglePopover(st, button, build) {
    var wasOpen = st.open && st.open.button === button;
    closePopover(st);
    if (wasOpen) return;

    var pop = build();
    st.bar.appendChild(pop);
    var left = button.offsetLeft;
    var maxLeft = st.bar.clientWidth - pop.offsetWidth - 2;
    pop.style.left = Math.max(0, Math.min(left, maxLeft)) + "px";
    button.setAttribute("aria-expanded", "true");
    pop.addEventListener("keydown", function (e) {
      if (e.key === "Escape") { e.preventDefault(); closePopover(st); restore(st); }
    });
    st.open = { pop: pop, button: button };
    if (pop._focusTarget) pop._focusTarget.focus();
  }

  function closePopover(st) {
    if (!st.open) return;
    st.open.button.setAttribute("aria-expanded", "false");
    if (st.open.pop.parentNode) st.open.pop.parentNode.removeChild(st.open.pop);
    st.open = null;
  }

  // ────────── Toolbar: selection and commands ────────────────

  /** Put focus and the saved selection back in the editor. */
  function restore(st) {
    var sel = window.getSelection();
    st.el.focus();
    if (st.range && st.el.contains(st.range.commonAncestorContainer)) {
      sel.removeAllRanges();
      sel.addRange(st.range);
    } else if (!st.el.contains(sel.anchorNode)) {
      var r = document.createRange();
      r.selectNodeContents(st.el);
      r.collapse(false);
      sel.removeAllRanges();
      sel.addRange(r);
    }
  }

  function exec(st, cmd, value, useCss) {
    restore(st);
    try { document.execCommand("styleWithCSS", false, useCss); } catch (e) { /* not supported */ }
    document.execCommand(cmd, false, value);
    changed(st);
    refresh(st);
  }

  // Code that saves the editor's HTML listens for "input". execCommand fires
  // it already, but direct DOM edits (updating an existing link) do not.
  function changed(st) {
    var ev;
    try { ev = new Event("input", { bubbles: true }); }
    catch (e) { ev = document.createEvent("Event"); ev.initEvent("input", true, false); }
    st.el.dispatchEvent(ev);
  }

  function refresh(st) {
    for (var name in st.buttons) {
      if (!BUTTONS[name].state) continue;
      var on = false;
      try { on = document.queryCommandState(BUTTONS[name].cmd); } catch (e) { /* ignore */ }
      st.buttons[name].classList.toggle("is-active", on);
      st.buttons[name].setAttribute("aria-pressed", on ? "true" : "false");
    }
  }

  function supported(cmd) {
    try { return document.queryCommandSupported(cmd); } catch (e) { return false; }
  }

  function escapeAttr(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  }

  function toHex(css) {
    var m = /^rgba?\((\d+),\s*(\d+),\s*(\d+)/i.exec(css || "");
    if (!m) return /^#[0-9a-f]{6}$/i.test(css) ? css : null;
    return "#" + [m[1], m[2], m[3]].map(function (n) {
      return ("0" + parseInt(n, 10).toString(16)).slice(-2);
    }).join("");
  }

  // ────────── Public API ─────────────────────────────────────

  /** Get the editor's HTML content. */
  WordPasteEditor.prototype.getHTML = function () {
    return this._editor.innerHTML;
  };

  /** Set the editor's HTML content. */
  WordPasteEditor.prototype.setHTML = function (html) {
    this._editor.innerHTML = html;
  };

  /** Remove all event listeners and clean up. */
  WordPasteEditor.prototype.destroy = function () {
    this._editor.removeEventListener("paste", this._handlers.paste);
    this._editor.removeEventListener("drop", this._handlers.drop);
    this._editor.removeEventListener("dragover", this._handlers.dragover);
    this._editor.removeEventListener("keydown", this._handlers.keydown);
    if (this._toggle && this._handlers.toggle) {
      this._toggle.removeEventListener("click", this._handlers.toggle);
    }
    if (this._toolbar) {
      destroyToolbar(this._toolbar);
      this._toolbar = null;
    }
  };

  /** Expose WordCleaner as a static property for convenience. */
  WordPasteEditor.WordCleaner = WordCleaner;

  return WordPasteEditor;
});
