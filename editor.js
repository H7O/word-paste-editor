/**
 * WordPasteEditor – minimal contenteditable HTML editor with Word-paste support.
 * Zero dependencies (uses WordCleaner from word-cleaner.js).
 *
 * @param {string|HTMLElement} element - The contenteditable element or CSS selector.
 * @param {Object} [options]
 * @param {string|HTMLElement} [options.sourceView] - Textarea element or selector for HTML source view.
 * @param {string|HTMLElement} [options.sourceToggle] - Button element or selector to toggle source view.
 * @param {string} [options.placeholder] - Placeholder text for the editor.
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
  };

  /** Expose WordCleaner as a static property for convenience. */
  WordPasteEditor.WordCleaner = WordCleaner;

  return WordPasteEditor;
});
