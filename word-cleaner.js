/**
 * Word HTML Cleaner — zero-dependency module that sanitises HTML pasted from
 * Microsoft Word (and similar rich-text sources) while preserving meaningful
 * formatting: tables, cell colours, fonts, bold/italic/underline, text colour,
 * alignment, lists, and images.
 *
 * Usage:
 *   const clean = WordCleaner.clean(dirtyHtml);
 */
(function (root, factory) {
  if (typeof define === "function" && define.amd) {
    define([], factory);
  } else if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    root.WordCleaner = factory();
  }
}(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  // ───────────────────────── helpers ──────────────────────────

  /** Parse an HTML string into a DOM document. */
  function parseHTML(html) {
    return new DOMParser().parseFromString(html, "text/html");
  }

  /** Serialise a DOM node's inner HTML back to a string. */
  function innerHTML(node) {
    var div = document.createElement("div");
    var childNodes = node.childNodes;
    for (var i = 0; i < childNodes.length; i++) {
      div.appendChild(childNodes[i].cloneNode(true));
    }
    return div.innerHTML;
  }

  // ───────── detect whether HTML originated from Word ────────

  function isWordHTML(html) {
    return (
      /class="?Mso/i.test(html) ||
      /xmlns:o=/.test(html) ||
      /xmlns:w=/.test(html) ||
      /urn:schemas-microsoft-com:office/.test(html) ||
      /mso-/i.test(html) ||
      /<o:p>/i.test(html)
    );
  }

  // ────────── CSS property white-list / mapping ──────────────

  /** Standard CSS properties we keep when they carry useful formatting. */
  var KEEP_PROPS = [
    "color",
    "background-color",
    "background",
    "font-family",
    "font-size",
    "font-weight",
    "font-style",
    "text-decoration",
    "text-decoration-line",
    "text-align",
    "vertical-align",
    "border",
    "border-top",
    "border-right",
    "border-bottom",
    "border-left",
    "border-collapse",
    "border-color",
    "border-width",
    "border-style",
    "border-top-color",
    "border-right-color",
    "border-bottom-color",
    "border-left-color",
    "border-top-width",
    "border-right-width",
    "border-bottom-width",
    "border-left-width",
    "border-top-style",
    "border-right-style",
    "border-bottom-style",
    "border-left-style",
    "width",
    "height",
    "min-width",
    "min-height",
    "padding",
    "padding-top",
    "padding-right",
    "padding-bottom",
    "padding-left",
    "margin-bottom",
    "margin-top",
    "list-style-type",
    "text-indent",
    "line-height",
    "white-space",
  ];

  var KEEP_SET = {};
  KEEP_PROPS.forEach(function (p) {
    KEEP_SET[p] = true;
  });

  // ──────────── MSO → standard CSS conversions ───────────────

  /**
   * Attempt to convert some `mso-*` properties into their CSS equivalents.
   * Returns an object of {property: value} pairs to merge in.
   */
  function convertMsoProperties(cssText) {
    var extra = {};

    // mso-shading / mso-pattern – cell / paragraph background
    var shading = cssText.match(
      /(?:mso-shading|mso-pattern)[^;]*windowtext;\s*fill\s*:\s*([^;"]+)/i
    );
    if (shading) {
      extra["background-color"] = shading[1].trim();
    }

    // mso-highlight
    var highlight = cssText.match(/mso-highlight\s*:\s*([^;"]+)/i);
    if (highlight) {
      extra["background-color"] = highlight[1].trim();
    }

    // background (sometimes Word uses background: #fff directly but browser ignores it
    // because it's mixed with mso props)
    var bg = cssText.match(
      /(?:^|;)\s*background\s*:\s*(#[0-9a-f]{3,8}|rgba?\([^)]+\)|[a-z]+)\s*(?:;|$)/i
    );
    if (bg) {
      extra["background-color"] = bg[1].trim();
    }

    // mso-border-*-alt  (Word sometimes puts real border info here)
    var borderAltRe =
      /mso-border-(top|right|bottom|left)-alt\s*:\s*([^;"]+)/gi;
    var m;
    while ((m = borderAltRe.exec(cssText)) !== null) {
      extra["border-" + m[1]] = m[2]
        .replace(/\s+\d+(\.\d+)?pt/g, function (v) {
          return v; // keep pt values as-is, browser can handle them
        })
        .trim();
    }

    return extra;
  }

  // ──────────── inline style cleaning ────────────────────────

  /**
   * Parse an inline style string, keep only white-listed properties plus any
   * MSO-converted properties, and return a cleaned style string.
   */
  function cleanStyle(styleText) {
    if (!styleText) return "";

    // First extract any MSO conversions from the raw text
    var msoExtras = convertMsoProperties(styleText);

    // Parse the style text into property-value pairs
    var declarations = styleText.split(";");
    var kept = {};

    for (var i = 0; i < declarations.length; i++) {
      var decl = declarations[i].trim();
      if (!decl) continue;
      var colon = decl.indexOf(":");
      if (colon < 0) continue;
      var prop = decl.substring(0, colon).trim().toLowerCase();
      var val = decl.substring(colon + 1).trim();

      // Skip mso-* and other junk
      if (/^mso-/.test(prop)) continue;
      if (/^-ms-/.test(prop)) continue;
      if (prop === "tab-stops") continue;
      if (prop === "layout-grid-mode") continue;
      if (prop === "text-autospace") continue;
      if (prop === "word-break") continue; // Word-inserted

      if (KEEP_SET[prop]) {
        kept[prop] = val;
      }
    }

    // Merge MSO-converted extras (they take lower precedence — don't overwrite
    // explicit values already found).
    for (var key in msoExtras) {
      if (!kept[key]) {
        kept[key] = msoExtras[key];
      }
    }

    // Remove "background-color: transparent" and similar no-ops
    if (
      kept["background-color"] &&
      /^(transparent|inherit|initial)$/i.test(kept["background-color"])
    ) {
      delete kept["background-color"];
    }
    if (
      kept["background"] &&
      /^(transparent|inherit|initial)$/i.test(kept["background"])
    ) {
      delete kept["background"];
    }

    // Convert font-family: strip quotes around generic families, keep first 2
    if (kept["font-family"]) {
      kept["font-family"] = cleanFontFamily(kept["font-family"]);
    }

    // Remove "color: windowtext" (Word's default)
    if (kept["color"] && /^windowtext$/i.test(kept["color"])) {
      delete kept["color"];
    }

    // Build the result string
    var parts = [];
    for (var p in kept) {
      parts.push(p + ": " + kept[p]);
    }
    return parts.join("; ");
  }

  /** Simplify font-family stacks — keep up to 2 families. */
  function cleanFontFamily(val) {
    var fonts = val.split(",").map(function (f) {
      return f.trim().replace(/^["']|["']$/g, "");
    });

    // Remove Word internal fonts
    fonts = fonts.filter(function (f) {
      return !/^(Calibri Light|Aptos|Aptos Display)$/i.test(f) === false
        ? true
        : true;
    });

    if (fonts.length > 3) fonts = fonts.slice(0, 3);
    return fonts
      .map(function (f) {
        return /\s/.test(f) ? '"' + f + '"' : f;
      })
      .join(", ");
  }

  // ─────────────── Word list (mso-list) handling ─────────────

  /**
   * Word doesn't use real <ul>/<ol>. Instead it uses:
   *   <p class=MsoListParagraph style='mso-list:l0 level1 lfo1'>
   *     <![if !supportLists]><span>1.<span>&nbsp;</span></span><![endif]>
   *     Item text
   *   </p>
   *
   * The marker between the supportLists comments is the number or bullet
   * Word displayed. preCleanHTML wraps it in <span class="__list-marker">.
   *
   * Numbered items keep that marker as literal text at the start of the
   * paragraph, so the numbers match Word exactly. That covers numbers like
   * "2.1" or "(a)", numbering that continues after other paragraphs, and
   * serial-number columns in tables, where every cell is a one-item list.
   * Consecutive bulleted items are grouped into proper <ul> lists.
   *
   * @param {Element} doc - Root to search (the parsed body).
   * @param {Object} listDefs - From parseListDefinitions: "l0:1" → true when
   *   that list level is bulleted.
   */
  function convertLists(doc, listDefs) {
    var paras = doc.querySelectorAll("p, h1, h2, h3, h4, h5, h6");
    var bullets = [];

    for (var i = 0; i < paras.length; i++) {
      var p = paras[i];
      var marker = findListMarker(p);
      if (!marker) continue;

      var listMatch = (p.getAttribute("style") || "").match(
        /mso-list\s*:\s*(l\d+)\s+level(\d+)/i
      );
      var level = listMatch ? parseInt(listMatch[2], 10) : 1;
      var text = (marker.textContent || "").replace(/[\s ]+/g, " ").trim();
      var defKey = listMatch ? listMatch[1].toLowerCase() + ":" + level : "";
      var isBullet = listDefs.hasOwnProperty(defKey)
        ? listDefs[defKey]
        : isBulletMarker(text, marker);

      var markerStyle = findMarkerStyle(marker);
      marker.parentNode.removeChild(marker);

      if (isBullet) {
        bullets.push({ element: p, level: level, content: p.innerHTML });
        continue;
      }

      dropHangingIndent(p, level);
      if (!text) continue; // numbering format "none"

      var ownerDoc = p.ownerDocument;
      var number = ownerDoc.createTextNode(text);
      if (markerStyle) {
        var span = ownerDoc.createElement("span");
        span.setAttribute("style", markerStyle);
        span.appendChild(number);
        number = span;
      }

      if (hasContent(p)) {
        p.insertBefore(ownerDoc.createTextNode(" "), p.firstChild);
      } else {
        // Only Word's empty-paragraph filler is left (as in a serial-number
        // cell), and it would push the number off centre.
        while (p.firstChild) p.removeChild(p.firstChild);
      }
      p.insertBefore(number, p.firstChild);
    }

    // Markers outside a paragraph are padding and stray numbers.
    var leftovers = doc.querySelectorAll("span.__list-marker");
    for (var k = 0; k < leftovers.length; k++) {
      leftovers[k].parentNode.removeChild(leftovers[k]);
    }

    // Group consecutive bulleted paragraphs into lists
    var groups = [];
    var currentGroup = null;
    for (var j = 0; j < bullets.length; j++) {
      var prev = j > 0 ? bullets[j - 1] : null;
      if (!currentGroup || prev.element.nextElementSibling !== bullets[j].element) {
        currentGroup = [];
        groups.push(currentGroup);
      }
      currentGroup.push(bullets[j]);
    }

    for (var g = 0; g < groups.length; g++) {
      var group = groups[g];
      var firstPara = group[0].element;
      firstPara.parentNode.insertBefore(buildNestedList(group), firstPara);
      for (var m = 0; m < group.length; m++) {
        group[m].element.parentNode.removeChild(group[m].element);
      }
    }
  }

  /**
   * Read Word's list definitions from the <style> block, for example
   *   @list l0:level1 {mso-level-number-format:bullet; ...}
   * Returns "l0:1" → true for bulleted levels, false for numbered ones.
   */
  function parseListDefinitions(html) {
    var defs = {};
    var re = /@list\s+(l\d+):level(\d+)\s*\{([^}]*)\}/gi;
    var m;
    while ((m = re.exec(html)) !== null) {
      defs[m[1].toLowerCase() + ":" + m[2]] =
        /mso-level-number-format\s*:\s*(bullet|image)/i.test(m[3]);
    }
    return defs;
  }

  /**
   * Find the element holding a paragraph's list marker: the span preCleanHTML
   * made from the supportLists comments, or, when the comments were already
   * stripped, Word's <span style="mso-list:Ignore"> and the spans that only
   * wrap it.
   */
  function findListMarker(p) {
    var marker = p.querySelector("span.__list-marker");
    if (marker) return marker;

    var spans = p.getElementsByTagName("span");
    for (var i = 0; i < spans.length; i++) {
      if (/mso-list\s*:\s*ignore/i.test(spans[i].getAttribute("style") || "")) {
        marker = spans[i];
        while (
          marker.parentNode !== p &&
          marker.parentNode.tagName === "SPAN" &&
          marker.parentNode.textContent === marker.textContent
        ) {
          marker = marker.parentNode;
        }
        return marker;
      }
    }
    return null;
  }

  /**
   * The style of the first element in the marker that carries formatting
   * (font, size, colour), so a literal number looks the way Word showed it.
   * Word's padding spans only set the "font" shorthand, which cleanStyle drops.
   */
  function findMarkerStyle(marker) {
    var els = [marker].concat(
      Array.prototype.slice.call(marker.getElementsByTagName("*"))
    );
    for (var i = 0; i < els.length; i++) {
      var style = els[i].getAttribute("style");
      if (style && cleanStyle(style)) return style;
    }
    return "";
  }

  /**
   * Fallback when the list definitions are missing: a bullet is a single
   * bullet-like character, or anything set in a symbol font.
   */
  function isBulletMarker(text, marker) {
    if (/font-family\s*:\s*["']?(Symbol|Wingdings|Webdings)/i.test(marker.outerHTML)) {
      return true;
    }
    // · • ‣ ◦ ▪ ▫ ■ □ ● ○ ◆ ◇ ► ▸ ➢ ➤ ✓ ✔ § Ø ü o * – — -
    return /^[·•‣◦▪▫■□●○◆◇►▸➢➤✓✔§Øüo*–—-]$/.test(text);
  }

  /** True when the paragraph has text or an image, table or line break. */
  function hasContent(p) {
    return (
      (p.textContent || "").replace(/[\s ]+/g, "") !== "" ||
      !!p.querySelector("img, table, br")
    );
  }

  /**
   * Word lays a list paragraph out as a hanging indent: a left margin with a
   * negative text-indent, sized for the number and a tab. With the number
   * inline that only shifts the first line (in a table cell, out of the
   * cell). Drop it, and indent nested levels a little instead.
   */
  function dropHangingIndent(p, level) {
    var style = (p.getAttribute("style") || "").replace(
      /(^|;)\s*(text-indent|margin-left|mso-text-indent-alt)\s*:[^;]*/gi,
      "$1"
    );
    if (level > 1) style += ";padding-left:" + (level - 1) * 0.25 + "in";
    p.setAttribute("style", style);
  }

  /** Build a (possibly nested) <ul> from a group of bulleted items. */
  function buildNestedList(items) {
    var ownerDoc = items[0].element.ownerDocument;
    var rootList = ownerDoc.createElement("ul");
    var stack = [{ list: rootList, level: items[0].level }];

    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var li = ownerDoc.createElement("li");
      li.innerHTML = item.content;

      while (stack.length > 1 && stack[stack.length - 1].level >= item.level) {
        stack.pop();
      }

      if (item.level > stack[stack.length - 1].level) {
        // Need to nest deeper
        var subList = ownerDoc.createElement("ul");
        var parentList = stack[stack.length - 1].list;
        var lastLi = parentList.lastElementChild;
        if (!lastLi) {
          lastLi = ownerDoc.createElement("li");
          parentList.appendChild(lastLi);
        }
        lastLi.appendChild(subList);
        stack.push({ list: subList, level: item.level });
      }

      stack[stack.length - 1].list.appendChild(li);
    }

    return rootList;
  }

  // ───────────── node-level cleaning (recursive) ─────────────

  /** Tags we remove entirely (including children). */
  var REMOVE_TAGS = {
    STYLE: true,
    SCRIPT: true,
    META: true,
    LINK: true,
    TITLE: true,
    XML: true,
    "O:P": true, // Word namespace tags treated as elements
    "W:WORDDOCUMENT": true,
    "W:VIEW": true,
    "W:ZOOM": true,
    "W:TRACKMOVES": true,
    "W:TRACKFORMATTING": true,
    "W:PUNCTUATIONKERNING": true,
    "W:DRAWINGGRIDHORIZONTALSPACING": true,
    "W:DISPLAYHORIZONTALDRAWINGGRIDEACHTICK": true,
    "W:COMPATIBILITY": true,
    "W:BREAKWRAPPEDTABLES": true,
    "W:SNAPTOGRIDINCELL": true,
    "W:WRAPTEXTWITHPUNCT": true,
    "W:USELOCALIZEDBUILTINSTYLES": true,
    "W:LATENTSTYLES": true,
    "W:LSTYLE": true,
  };

  /** Tags we unwrap (keep children, remove the tag itself). */
  var UNWRAP_TAGS = {
    "O:P": true,
    FONT: false, // handled separately so we can migrate color/face
  };

  /** Tags we always allow. */
  var ALLOW_TAGS = {
    P: true,
    BR: true,
    B: true,
    STRONG: true,
    I: true,
    EM: true,
    U: true,
    S: true,
    STRIKE: true,
    SUB: true,
    SUP: true,
    SPAN: true,
    DIV: true,
    A: true,
    TABLE: true,
    THEAD: true,
    TBODY: true,
    TFOOT: true,
    TR: true,
    TD: true,
    TH: true,
    CAPTION: true,
    COL: true,
    COLGROUP: true,
    UL: true,
    OL: true,
    LI: true,
    H1: true,
    H2: true,
    H3: true,
    H4: true,
    H5: true,
    H6: true,
    BLOCKQUOTE: true,
    PRE: true,
    CODE: true,
    HR: true,
    IMG: true,
  };

  /** Attributes that are safe to keep. */
  var ALLOW_ATTRS = {
    style: true,
    href: true,
    src: true,
    alt: true,
    title: true,
    colspan: true,
    rowspan: true,
    width: true,
    height: true,
    align: true,
    valign: true,
    "border": true,
    cellpadding: true,
    cellspacing: true,
    scope: true,
  };

  /**
   * Walk the DOM tree, cleaning each node in-place.
   */
  function cleanNode(node) {
    if (node.nodeType === 3) return; // text node – keep as-is
    if (node.nodeType === 8) {
      // comment – remove
      node.parentNode && node.parentNode.removeChild(node);
      return;
    }
    if (node.nodeType !== 1) {
      node.parentNode && node.parentNode.removeChild(node);
      return;
    }

    var tag = node.tagName.toUpperCase();

    // 1. Remove tags we don't want at all
    if (REMOVE_TAGS[tag]) {
      node.parentNode && node.parentNode.removeChild(node);
      return;
    }

    // 2. Handle <o:p> – typically just wraps &nbsp;, unwrap it
    if (tag === "O:P") {
      unwrapNode(node);
      return;
    }

    // 3. Handle namespace tags (v:*, w:*, etc.) – remove them
    if (tag.indexOf(":") !== -1) {
      node.parentNode && node.parentNode.removeChild(node);
      return;
    }

    // 4. Handle <font> — migrate to <span> with style
    if (tag === "FONT") {
      convertFontToSpan(node);
      // after conversion the node is now a <span>, continue cleaning children
      tag = "SPAN";
    }

    // 5. If tag not in allow-list, unwrap
    if (!ALLOW_TAGS[tag]) {
      unwrapNode(node);
      return;
    }

    // 6. Clean attributes
    cleanAttributes(node);

    // 7. Recurse children (iterate backwards because we may remove/unwrap)
    var children = Array.prototype.slice.call(node.childNodes);
    for (var i = 0; i < children.length; i++) {
      cleanNode(children[i]);
    }

    // 8. Remove empty spans/divs that contribute nothing
    if (
      (tag === "SPAN" || tag === "DIV") &&
      !node.hasAttribute("style") &&
      !node.hasChildNodes()
    ) {
      node.parentNode && node.parentNode.removeChild(node);
      return;
    }

    // 9. Unwrap spans that have no style (they're just noise)
    if (tag === "SPAN" && !node.getAttribute("style")) {
      unwrapNode(node);
    }
  }

  /** Move all children of `node` before it, then remove it. */
  function unwrapNode(node) {
    var parent = node.parentNode;
    if (!parent) return;
    while (node.firstChild) {
      parent.insertBefore(node.firstChild, node);
    }
    parent.removeChild(node);
  }

  /** Convert <font color=X face=Y size=Z> into <span style="...">. */
  function convertFontToSpan(font) {
    var span = document.createElement("span");
    var style = [];

    var color = font.getAttribute("color");
    if (color) style.push("color: " + color);

    var face = font.getAttribute("face");
    if (face) style.push("font-family: " + face);

    var size = font.getAttribute("size");
    if (size) {
      var sizeMap = { 1: "8pt", 2: "10pt", 3: "12pt", 4: "14pt", 5: "18pt", 6: "24pt", 7: "36pt" };
      if (sizeMap[size]) style.push("font-size: " + sizeMap[size]);
    }

    if (style.length) span.setAttribute("style", style.join("; "));

    // Move children
    while (font.firstChild) {
      span.appendChild(font.firstChild);
    }

    font.parentNode && font.parentNode.replaceChild(span, font);
  }

  /** Remove disallowed attributes; clean the style attribute. */
  function cleanAttributes(node) {
    var attrs = Array.prototype.slice.call(node.attributes);
    for (var i = 0; i < attrs.length; i++) {
      var name = attrs[i].name.toLowerCase();
      if (name === "class") {
        // Remove all Word classes
        node.removeAttribute("class");
      } else if (name === "style") {
        var cleaned = cleanStyle(attrs[i].value);
        if (cleaned) {
          node.setAttribute("style", cleaned);
        } else {
          node.removeAttribute("style");
        }
      } else if (!ALLOW_ATTRS[name]) {
        node.removeAttribute(attrs[i].name);
      }
    }
  }

  // ────────────── high-level text cleaning ────────────────────

  /** Pre-process raw HTML string before DOM parsing. */
  function preCleanHTML(html) {
    // Preserve list markers — we'll handle them in convertLists. Word's
    // clipboard writes them as <![if !supportLists]>…<![endif]>, saved and
    // emailed HTML as <!--[if !supportLists]-->…<!--[endif]-->.
    html = html.replace(/<!(?:--)?\[if\s+!supportLists\](?:--)?>([\s\S]*?)<!(?:--)?\[endif\](?:--)?>/gi, function(match, inner) {
      return '<span class="__list-marker">' + inner + '</span>';
    });

    // Remove conditional comments: <!--[if ...]>...<![endif]-->
    html = html.replace(/<!--\[if[\s\S]*?\]>[\s\S]*?<!\[endif\]-->/gi, "");
    html = html.replace(/<!--[\s\S]*?-->/g, "");

    // Remove XML declarations and processing instructions
    html = html.replace(/<\?xml[\s\S]*?\?>/gi, "");

    // Remove Word-specific XML namespace declarations in the root tag
    // (these sometimes cause DOMParser issues)
    html = html.replace(/<html[^>]*>/gi, "<html>");

    // Remove <head>…</head> entirely
    html = html.replace(/<head[\s\S]*?<\/head>/gi, "");

    return html;
  }

  /** Post-process: collapse excessive whitespace / empty paragraphs. */
  function postClean(doc) {
    // Remove empty paragraphs that are just &nbsp;
    var paras = doc.querySelectorAll("p");
    for (var i = 0; i < paras.length; i++) {
      var text = paras[i].textContent || "";
      if (
        !text.trim() &&
        !paras[i].querySelector("img, table, br") &&
        paras[i].innerHTML.replace(/&nbsp;/gi, "").trim() === ""
      ) {
        // Keep one <br> equivalent to avoid losing spacing
        // but only if surrounded by other content
        var prev = paras[i].previousElementSibling;
        var next = paras[i].nextElementSibling;
        if (prev && next) {
          // Replace with a <br>
          var br = doc.ownerDocument.createElement("br");
          paras[i].parentNode.replaceChild(br, paras[i]);
        } else {
          paras[i].parentNode && paras[i].parentNode.removeChild(paras[i]);
        }
      }
    }
  }

  // ────────────── main clean function ────────────────────────

  /**
   * Clean pasted HTML. If the HTML doesn't appear to be from Word, a lighter
   * sanitisation is applied (just stripping scripts/styles and dangerous attrs).
   *
   * @param {string} html - Raw HTML from clipboard.
   * @returns {string} Cleaned HTML string.
   */
  function clean(html) {
    if (!html) return "";

    var fromWord = isWordHTML(html);
    var listDefs = {};

    if (fromWord) {
      // Read before preCleanHTML removes the <head> they live in
      listDefs = parseListDefinitions(html);
      html = preCleanHTML(html);
    }

    var doc = parseHTML(html);

    if (fromWord) {
      convertLists(doc.body, listDefs);
    }

    // Walk and clean all nodes
    var children = Array.prototype.slice.call(doc.body.childNodes);
    for (var i = 0; i < children.length; i++) {
      cleanNode(children[i]);
    }

    if (fromWord) {
      postClean(doc.body);
    }

    var result = doc.body.innerHTML;

    // Final string-level cleanup
    // Collapse runs of <br>
    result = result.replace(/(<br\s*\/?\s*>){3,}/gi, "<br><br>");

    // Remove completely empty tags at the top level
    result = result.replace(/<(p|div|span)>\s*<\/\1>/gi, "");

    // Trim
    result = result.trim();

    return result;
  }

  // ────────────── public API ─────────────────────────────────

  return {
    clean: clean,
    isWordHTML: isWordHTML,
  };
}));
