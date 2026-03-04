/**
 * Editor – minimal contenteditable HTML editor with Word-paste support.
 * Zero dependencies (uses WordCleaner from word-cleaner.js).
 */
(function () {
  "use strict";

  var editor = document.getElementById("editor");
  var sourceView = document.getElementById("sourceView");
  var modeToggle = document.getElementById("modeToggle");
  var isSourceMode = false;

  // ────────── Paste handling ──────────────────────────────────

  editor.addEventListener("paste", function (e) {
    // Grab the HTML flavour from the clipboard
    var clipboardData = e.clipboardData || window.clipboardData;
    if (!clipboardData) return;

    var html = clipboardData.getData("text/html");
    var plainText = clipboardData.getData("text/plain");

    // If there's HTML content, clean it and insert
    if (html) {
      e.preventDefault();

      var cleaned = WordCleaner.clean(html);

      // Insert the cleaned HTML at the current cursor position
      insertHTML(cleaned);
      return;
    }

    // If only plain text, let the browser handle it naturally
    // (or we could manually insert to avoid <div> wrapping quirks)
    if (plainText && !html) {
      e.preventDefault();
      // Convert line breaks to <br> and insert
      var escaped = escapeHTML(plainText);
      escaped = escaped.replace(/\n/g, "<br>");
      insertHTML(escaped);
    }
  });

  // ────────── Insertion helper ────────────────────────────────

  /** Insert HTML at the current selection/cursor position in the editor. */
  function insertHTML(html) {
    // Modern approach: insertHTML command
    if (document.queryCommandSupported && document.queryCommandSupported("insertHTML")) {
      document.execCommand("insertHTML", false, html);
      return;
    }

    // Fallback: manual range insertion
    var sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) {
      // No selection, just append
      editor.innerHTML += html;
      return;
    }

    var range = sel.getRangeAt(0);
    range.deleteContents();

    var frag = document.createRange().createContextualFragment(html);
    var lastNode = frag.lastChild;
    range.insertNode(frag);

    // Move cursor to end of inserted content
    if (lastNode) {
      var newRange = document.createRange();
      newRange.setStartAfter(lastNode);
      newRange.collapse(true);
      sel.removeAllRanges();
      sel.addRange(newRange);
    }
  }

  /** Escape HTML special chars for plain-text insertion. */
  function escapeHTML(text) {
    var div = document.createElement("div");
    div.appendChild(document.createTextNode(text));
    return div.innerHTML;
  }

  // ────────── Source view toggle ──────────────────────────────

  modeToggle.addEventListener("click", function () {
    isSourceMode = !isSourceMode;

    if (isSourceMode) {
      // Switch to source view
      sourceView.value = formatHTML(editor.innerHTML);
      editor.style.display = "none";
      sourceView.style.display = "block";
      modeToggle.textContent = "View Editor";
      sourceView.focus();
    } else {
      // Switch back to editor
      editor.innerHTML = sourceView.value;
      sourceView.style.display = "none";
      editor.style.display = "block";
      modeToggle.textContent = "View Source";
      editor.focus();
    }
  });

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

  // ────────── Drop handling (treat like paste) ───────────────

  editor.addEventListener("drop", function (e) {
    var dt = e.dataTransfer;
    if (!dt) return;

    var html = dt.getData("text/html");
    if (html) {
      e.preventDefault();
      var cleaned = WordCleaner.clean(html);
      // Place cursor at drop position
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

      insertHTML(cleaned);
    }
  });

  // Prevent default dragover to allow drop
  editor.addEventListener("dragover", function (e) {
    e.preventDefault();
  });

  // ────────── Keyboard shortcut: Ctrl+Shift+V = paste as plain text ──

  editor.addEventListener("keydown", function (e) {
    // Ctrl+Shift+V — the browser may already handle this, but just in case
    if (e.ctrlKey && e.shiftKey && e.key === "V") {
      // Let browser handle paste-as-plain-text
      return;
    }

    // Tab key — insert spaces instead of moving focus
    if (e.key === "Tab") {
      e.preventDefault();
      insertHTML("&nbsp;&nbsp;&nbsp;&nbsp;");
    }
  });

  // ────────── Init ───────────────────────────────────────────

  editor.focus();
})();
