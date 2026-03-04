# word-paste-editor

A zero-dependency, toolbar-free HTML editor that properly handles paste from Microsoft Word — preserving tables, cell colors, fonts, formatting, and more.

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

## Files

| File | Description |
|------|-------------|
| `editor.html` | Main page — contenteditable editor with minimal CSS |
| `word-cleaner.js` | Word HTML cleaner/sanitizer module |
| `editor.js` | Editor logic — paste interception, drag-drop, source view |

## Usage

1. Open `editor.html` in any modern browser
2. Copy content from Microsoft Word (tables, formatted text, lists, etc.)
3. Paste into the editor — formatting is preserved, Word junk is removed
4. Click **View Source** to see or edit the cleaned HTML

### Integration

To use in your own project, include both JS files and set up a `contenteditable` div:

```html
<div id="editor" contenteditable="true"></div>

<script src="word-cleaner.js"></script>
<script src="editor.js"></script>
```

Or use `WordCleaner` directly in your own paste handler:

```javascript
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

### `WordCleaner.clean(html)`

Takes a raw HTML string (e.g. from clipboard) and returns sanitized HTML with Word-specific markup removed and meaningful formatting preserved.

### `WordCleaner.isWordHTML(html)`

Returns `true` if the HTML appears to originate from Microsoft Word.

## Browser Support

Works in all modern browsers (Chrome, Firefox, Edge, Safari). Uses only standard Web APIs: `DOMParser`, `contenteditable`, `Selection`/`Range`, `document.execCommand`.

## License

MIT
