![License: MIT](https://img.shields.io/badge/license-MIT-blueviolet.svg)
![Static site](https://img.shields.io/badge/hosting-any%20static%20host-brightgreen.svg)
![Offline](https://img.shields.io/badge/works-offline-blue.svg)

# LibreSlides

> A presentation editor in the spirit of Google Slides and PowerPoint that runs entirely in your browser. Install it as an app and it works offline. No server, no account; your presentations stay on your device.

LibreSlides gives you a slide list, a canvas with drag/resize/rotate handles and snapping guides, themes, layouts, transitions, animations, speaker notes and a full-screen slideshow with presenter view. It opens and saves PowerPoint (`.pptx`) files. The whole app is a static site: build it once and host the `dist/` folder anywhere.

---

## Features

| Area | What you get |
|---|---|
| **Slides** | Slide list with thumbnails, multi-select, drag-to-reorder, duplicate, skip, cut/copy/paste; named sections that fold away, as in PowerPoint; 8 layouts (Title, Title and body, Section, Two columns, Title only, Big number, Quote, Blank); "Apply layout" keeps your text |
| **Canvas** | Select, move, resize and rotate with handles; marquee and Shift-click selection; smart guides that snap to the slide and other objects (hold Alt to move freely); arrow-key nudging; zoom (fit, 50–200 %, Ctrl + wheel) |
| **Objects** | Text boxes that grow as you type, shapes (with text inside, including all of PowerPoint's preset shapes and freeforms from its files), lines and arrows with arrowhead styles, pictures (insert, paste, drag & drop, replace, crop-to-fill, crop to a shape, greyscale/brightness/contrast, rounded corners, flip, alt text), tables (insert rows/columns, merge cells, header row, banded rows), charts (column, bar, line, area, pie, doughnut, scatter, with editable data), video and sound, groups (Ctrl+G), links to web pages or other slides (Ctrl+K), gradient and picture fills, shadows, slide comments |
| **Text** | Fonts (six bundled for offline use plus common system fonts), size, bold/italic/underline/strikethrough, colour, highlight, alignment, vertical alignment, line and paragraph spacing, bulleted and numbered lists with indent levels, capitals, letter spacing, superscript and subscript, text effects (shadow, glow, outline, hollow, gradient), text box margins, columns and vertical text, slide number and date fields, clear formatting. Formatting applies to selected text while typing, or to whole boxes when they're selected |
| **Design** | 9 themes with matching fonts, colours and graphics; theme colours follow theme changes; solid, gradient or picture backgrounds per slide or for all; 16:9, 16:10, 4:3, A4, Letter and custom sizes |
| **Arrange** | Bring forward/backward, to front/back, align and distribute, rotate and flip, lock position, exact size/position/rotation in Format options |
| **Motion** | 6 slide transitions (fade, slide, push, cover, zoom, flip); entrance, emphasis (pulse, spin, teeter, flash) and exit animations and motion paths, on click, with or after the previous one; preview in the editor |
| **Present** | Full-screen slideshow with keyboard, click and on-screen controls; black/white screen, jump to slide, speaker notes overlay. **Presenter view** opens the audience in a second window (drag it to the projector) and shows the current and next slide, notes and a timer here |
| **Files** | Open **PowerPoint** `.pptx` and **LibreSlides** `.lslides` files from your device (text and formatting, shapes and freeforms, pictures, lines, tables and table styles, charts, SmartArt (as editable shapes), video and audio, groups, links, animations, comments, backgrounds, notes, transitions, skipped slides, sections, embedded fonts saved by LibreOffice or LibreSlides; positions and styles inherited from layouts/masters are resolved). Everything above is written back to `.pptx` as native PowerPoint objects. In Chrome and Edge a presentation opened from a file stays linked to it: **Save** (Ctrl+S) writes back to the same file in its format, **Save as** (Ctrl+Shift+S) saves to a new file and keeps saving there. A title-bar indicator shows unsaved changes, and LibreSlides asks before closing with changes not yet saved to the file and warns if another app changed it. Other browsers download the file. PowerPoint files carry the bundled fonts they use (*File › Embed fonts in PowerPoint files*), so they look the same where those fonts aren't installed. You can also download a copy as `.pptx`, **PDF** (vector, via the print dialog) or `.lslides`, and drop files on the window to open them |
| **Storage** | Autosaved to IndexedDB, with a list of presentations (search, rename, copy, delete), backup/restore of everything, undo/redo, and protection against two tabs overwriting each other |
| **App** | Installable PWA (Install button on the start screen, or Help › Install) that works fully offline after the first visit, offers new versions with an *Update* button, has a "New presentation" shortcut on its icon, and registers as a handler for `.pptx` and `.lslides` files ("Open with LibreSlides") |

Keyboard shortcuts follow Slides and PowerPoint (Ctrl+M new slide, Ctrl+D duplicate, Ctrl+Enter / F5 present, Ctrl+Shift+> bigger text, Ctrl+↑/↓ to reorder…). Press **Ctrl+/** in the editor for the full list.

Not supported yet: real-time collaboration, WordArt's text warping (warped text comes in straight, keeping its effects), 3-D effects and SmartArt editing as a diagram (it's imported as shapes). Fonts PowerPoint itself embedded are compressed in a format LibreSlides can't read yet, so those fall back to similar fonts.

## Getting started

Requires Node.js 22.12 or newer.

```bash
npm install
npm run dev        # development server with hot reload
npm run build      # production build in dist/
npm run preview    # serve the production build locally
npm test           # unit tests (Vitest)
```

### Installing and using it offline

Open the hosted app (or `npm run build && npm run preview`) and click **Install app** on the start screen, or **Help › Install LibreSlides** in the editor. Where the browser can install directly (Chrome, Edge, Android) this shows its install dialog; elsewhere it explains the steps (Safari on macOS: *File › Add to Dock*; iOS: *Share › Add to Home Screen*; Firefox on computers can't install web apps but still works offline in a tab). Installed, LibreSlides gets its own window and launcher icon, and from then on works without a connection: the app, its fonts and PowerPoint import/export are all cached on the first visit.

### Saving to files on your device

Every change is autosaved inside LibreSlides (IndexedDB), so nothing is lost if you don't save to a file. To work with files:

- **Open** (Ctrl+O, the start screen, drag & drop, or *Open with → LibreSlides* in your file manager once installed) a `.pptx` or `.lslides` file.
- In Chrome, Edge and other Chromium browsers the presentation stays **linked to that file**: **Save** (Ctrl+S, or the file name in the title bar) writes your changes back to it in its own format. After a restart the browser may ask once to allow LibreSlides to edit the file. Opening the same file again goes back to the same presentation.
- **Save as** (Ctrl+Shift+S) saves to a new `.pptx` or `.lslides` file and links to it from then on.
- `.lslides` keeps everything. `.pptx` opens in PowerPoint, Keynote, Google Slides and LibreOffice, but doesn't keep LibreSlides animations or theme switching.
- Firefox and Safari can't write to files directly, so Save downloads a file instead.

Presentations kept inside LibreSlides live in the browser's storage for the app. Clearing site data deletes them, so save the ones you care about to files, or use **Back up all presentations** (the ⋯ menu next to *Recent presentations*) now and then.

### Deploying

`dist/` is plain static files with relative URLs, so it works from any host or sub-path: GitHub Pages, Netlify, Cloudflare Pages, S3, `python -m http.server` or a USB stick. The service worker needs HTTPS (or `localhost`) to make the app installable and offline.

The included workflow (`.github/workflows/deploy.yml`) runs the tests and publishes to **GitHub Pages** on every push to `main`. Enable it under *Settings → Pages → Build and deployment → Source: GitHub Actions*.

## Project structure

```
index.html                 App shell
src/
  main.js                  Router (#/ start, #/p/:id editor, #/show/:id audience window), file opening, PWA
  model/
    deck.js                Presentation model: slides, elements, layouts, pictures
    themes.js              Themes and theme colours
    shapes.js              Shape outlines (named after PowerPoint's preset shapes)
    templates.js           Starter presentations
    fonts.js               Font list
  render/render.js         Draws slides to DOM — used by the editor, thumbnails, slideshow and print
  editor/
    editor-screen.js       The editor: state, undo/redo, autosave, commands, shortcuts
    canvas.js              Selection, drag/resize/rotate, snapping, drawing, in-place text editing
    filmstrip.js           Slide list
    toolbar.js             Menu bar and contextual toolbar
    sidepanel.js           Design, transition, animation and format panels
    richtext.js, history.js
  present/present.js       Slideshow player, presenter view, previews
  io/
    pptx-export.js         PowerPoint writer (pptxgenjs, plus transitions/skipped slides)
    pptx-import.js         PowerPoint reader
    pptx-sections.js       Slide sections in .pptx files
    file-access.js         Opening, saving back to and saving as files on the device
    files.js, print.js, images.js, sanitize.js
  storage/db.js            IndexedDB
  ui/                      Start screen, dialogs, popovers, colour picker, icons, theme
tests/unit/                Model, rich text, slideshow and .pptx round-trip tests
```

In the app, `window.libreslides.screen` holds the open editor for poking around in devtools, and **Help › About** has details to copy into a bug report.

## Privacy

Presentations never leave your browser unless you download them. There is no backend, no analytics and no account.

## License

Released under the [MIT License](./LICENSE).
