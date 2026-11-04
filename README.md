# word-paste-editor

A zero-dependency HTML editor that properly handles paste from Microsoft Word — preserving tables, cell colors, fonts, formatting, and more. Toolbar-free by default, with an optional formatting toolbar.

![Vanilla JS](https://img.shields.io/badge/Vanilla-JS-F7DF1E?logo=javascript&logoColor=black)
![Zero Dependencies](https://img.shields.io/badge/Dependencies-0-brightgreen)
![License](https://img.shields.io/badge/License-MIT-green)

## Features

- **Zero dependencies** — no jQuery, no npm packages, no CDN links, no frameworks
- **No toolbar by default** — clean `contenteditable` surface, paste and go
- **Optional formatting toolbar** — bold, italic, underline, text and highlight colour, alignment, links and clear formatting (see [Toolbar](#toolbar))
- **Word paste support** — intelligently cleans Word's messy HTML while preserving:
  - Table structure, borders, and cell background colors
  - Font families, sizes, and colors
  - Bold, italic, underline, strikethrough
  - Text alignment
  - Numbered lists, with the numbers Word showed kept as literal text (`1.`, `2.1`, `(a)`), including auto-numbered serial-number columns in tables
  - Bulleted lists (converts Word's `mso-list` to proper `<ul>`)
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

<!-- Only if you turn on the toolbar -->
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/word-paste-editor@1/dist/toolbar.css">
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

Copy `dist/editor.min.js` and `dist/word-cleaner.min.js` into your project and include them via script tags. Add `dist/toolbar.css` if you turn on the toolbar.

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
| `toolbar` | `boolean \| string[]` | Show a formatting toolbar: `true` for every button, or an array of button names. Off by default. See [Toolbar](#toolbar) |
| `linkSchemes` | `string[]` | Link schemes the toolbar's link button accepts. Default `["http", "https"]` |

### Instance Methods

| Method | Returns | Description |
|--------|---------|-------------|
| `getHTML()` | `string` | Get the editor's HTML content |
| `setHTML(html)` | `void` | Set the editor's HTML content |
| `destroy()` | `void` | Remove all event listeners and the toolbar, and clean up |

### Toolbar

Pass `toolbar: true` and include `toolbar.css` (`dist/toolbar.css`, or `word-paste-editor/toolbar.css` from npm):

```html
<link rel="stylesheet" href="toolbar.css">
<div id="editor" contenteditable="true"></div>

<script src="word-cleaner.js"></script>
<script src="editor.js"></script>
<script>
  new WordPasteEditor('#editor', { toolbar: true });
</script>
```

The toolbar is inserted just before the editor element. To pick the buttons, pass their names in order, with `"|"` for a separator:

```js
new WordPasteEditor('#editor', {
  toolbar: ['bold', 'italic', 'underline', '|', 'link', 'unlink'],
  linkSchemes: ['https', 'mailto']
});
```

| Button | Does |
|--------|------|
| `bold`, `italic`, `underline` | Toggle `<b>`, `<i>`, `<u>` |
| `color`, `highlight` | Text and highlight colour from a 30-colour palette or the native colour picker, as `<span style="color / background-color">` |
| `alignLeft`, `alignCenter`, `alignRight`, `justify` | Set `text-align` on the paragraph |
| `link`, `unlink` | Insert, edit or remove a link |
| `clear` | Remove inline formatting |

Formatting uses `document.execCommand`, so every change fires the editor's `input` event, as typing does. Code that saves the editor's HTML on `input` keeps working. The link box accepts only the schemes in `linkSchemes`, so `javascript:` links can't be entered. It adds `https://` to a bare domain, and `mailto:` to a bare email address when `mailto` is allowed. Toolbar buttons are `type="button"` and the popover inputs have no `name`, so a surrounding `<form>` is never submitted or sent extra fields. The buttons are disabled while the source view is showing.

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
- `dist/toolbar.css` — toolbar styles

## Browser Support

Works in all modern browsers (Chrome, Firefox, Edge, Safari). Uses only standard Web APIs: `DOMParser`, `contenteditable`, `Selection`/`Range`, `document.execCommand`.

## License

[MIT](LICENSE) © Hussein Al Bayati
