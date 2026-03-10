# word-paste-editor

A zero-dependency, toolbar-free HTML editor that properly handles paste from Microsoft Word — preserving tables, cell colors, fonts, formatting, and more.

![Vanilla JS](https://img.shields.io/badge/Vanilla-JS-F7DF1E?logo=javascript&logoColor=black)
![Zero Dependencies](https://img.shields.io/badge/Dependencies-0-brightgreen)
![License](https://img.shields.io/badge/License-MIT-green)

## Features

- **Zero dependencies** — no jQuery, no npm packages, no CDN links, no frameworks
- **No toolbar** — clean `contenteditable` surface, paste and go
- **Word paste support** — intelligently cleans Word's messy HTML while preserving:
  - Table structure, borders, and cell background colors
  - Font families, sizes, and colors
  - Bold, italic, underline, strikethrough
  - Text alignment
  - Numbered and bulleted lists (converts Word's `mso-list` to proper `<ol>`/`<ul>`)
  - Images
- **Sanitizes HTML** — strips Word conditional comments, XML namespace tags (`<o:p>`, `<w:*>`, `<v:*>`), `mso-*` CSS properties, and disallowed attributes
- **Source view** — toggle to see/edit the clean HTML source
- **Drag-and-drop** — drop rich content with the same cleaning applied
- **Works with Google Docs** paste as well

## Installation

### CDN (quickest)

```html
<script src="https://cdn.jsdelivr.net/npm/word-paste-editor@1/dist/word-cleaner.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/word-paste-editor@1/dist/editor.min.js"></script>
```

### npm

```bash
npm install word-paste-editor
```

```js
// ES module
import WordPasteEditor from 'word-paste-editor';

// CommonJS
const WordPasteEditor = require('word-paste-editor');

// Cleaner only
import WordCleaner from 'word-paste-editor/cleaner';
const WordCleaner = require('word-paste-editor/cleaner');
```

### Manual

Copy `dist/editor.min.js` and `dist/word-cleaner.min.js` into your project and include them via script tags.

## Quick Start

### Script Tags

```html
<div id="editor" contenteditable="true"></div>
<textarea id="sourceView" style="display:none"></textarea>
<button id="modeToggle">View Source</button>

<script src="word-cleaner.js"></script>
<script src="editor.js"></script>
<script>
  new WordPasteEditor('#editor', {
    sourceView: '#sourceView',
    sourceToggle: '#modeToggle'
  });
</script>
```

### Minimal (no source view)

```html
<div id="editor" contenteditable="true"></div>

<script src="word-cleaner.js"></script>
<script src="editor.js"></script>
<script>
  new WordPasteEditor('#editor');
</script>
```

### Standalone Cleaner

Use `WordCleaner` directly in your own paste handler without the editor UI:

```js
element.addEventListener('paste', function(e) {
  var html = e.clipboardData.getData('text/html');
  if (html) {
    e.preventDefault();
    var cleaned = WordCleaner.clean(html);
    // insert cleaned HTML however you like
  }
});
```

## API

### `WordPasteEditor(element, options)`

Creates an editor instance on the given element.

| Option | Type | Description |
|--------|------|-------------|
| `sourceView` | `string \| HTMLElement` | Textarea element or selector for HTML source view |
| `sourceToggle` | `string \| HTMLElement` | Button element or selector to toggle source view |
| `placeholder` | `string` | Placeholder text for the editor |

### Instance Methods

| Method | Returns | Description |
|--------|---------|-------------|
| `getHTML()` | `string` | Get the editor's HTML content |
| `setHTML(html)` | `void` | Set the editor's HTML content |
| `destroy()` | `void` | Remove all event listeners and clean up |

### `WordCleaner.clean(html)`

Takes a raw HTML string (e.g. from clipboard) and returns sanitized HTML with Word-specific markup removed and meaningful formatting preserved.

### `WordCleaner.isWordHTML(html)`

Returns `true` if the HTML appears to originate from Microsoft Word.

## Demo

Open `editor.html` in a browser to see the editor in action — no build step required.

```bash
# Or use a local server:
npx serve .
```

## Building

To generate minified dist files:

```bash
npm install
npm run build
```

This produces:
- `dist/editor.js` — unminified copy
- `dist/editor.min.js` — minified
- `dist/word-cleaner.js` — unminified copy
- `dist/word-cleaner.min.js` — minified

## Browser Support

Works in all modern browsers (Chrome, Firefox, Edge, Safari). Uses only standard Web APIs: `DOMParser`, `contenteditable`, `Selection`/`Range`, `document.execCommand`.

## License

[MIT](LICENSE) © Hussein Al Bayati
