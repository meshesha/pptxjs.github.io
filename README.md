# pptx.js.org

Website of [PPTXjs](https://github.com/meshesha/PPTXjs), served by GitHub Pages at <https://pptx.js.org>.

Plain static HTML/CSS/JS with no build step. It uses the PPTXjs v2 browser bundle copied into `lib/`.

| Path | Content |
|---|---|
| `index.html` | Home page with a live viewer |
| `pages/demos.html` | Interactive demos: all slides, slideshow, your own file, editor, convert/export, code |
| `pages/docs.html` | Documentation |
| `pages/demo2.html`, `demo3.html`, `demo_revealjs.html` | Redirects for old v1 links |
| `pages/Sample_demo1.pptx`, `samples/` | Sample presentations |
| `assets/` | Site styles and scripts |
| `lib/` | PPTXjs v2 (`pptxjs.umd.js`, `pptxjs.css`, `VERSION`) |

## Update the library

After building PPTXjs (`npm run build` in the PPTXjs repository):

```bash
cp ../PPTXjs/dist/pptxjs.umd.js ../PPTXjs/dist/pptxjs.css lib/
```

Then update `lib/VERSION`.

## Preview locally

Any static server works, for example:

```bash
npx serve .
```

Opening the files directly via `file://` does not work, because browsers block `fetch()` of the sample decks.
