# dom2pptx

Turn HTML from the DOM into PowerPoint (`.pptx`) decks, by way of
[`@shbernal/ts-pptx`](https://www.npmjs.com/package/@shbernal/ts-pptx).

This library owns exactly one link in that chain: **HTML/DOM → `ts-pptx` calls**.
It reads a rendered document, builds an intermediate slide model (IR), and drives
the writer. It does not emit OOXML itself — that is `ts-pptx`'s job — and it does
not generate the HTML — that is the caller's job.

```
your app / AI agent               produces HTML slides
        │  imports
        ▼
dom2pptx (this package)           rendered DOM → IR → ts-pptx calls
        │  depends on
        ▼
@shbernal/ts-pptx                 IR-agnostic OOXML emitter (custGeom, etc.)
```

## Why

**1. Preview decks on the web before generating them.** HTML renders in a
browser instantly; PPTX does not. If the deck is authored as HTML, what you see
in the page is the preview, and this library is what makes that preview
redeemable as a real `.pptx`.

**2. Let AI agents build decks.** Models are good at emitting HTML and bad at
emitting OOXML. Give an agent an HTML target and it can produce a deck; this
library is the adapter that makes the HTML land as editable slides.

**3. (Deferred) Round-tripping.** Accept a `.pptx` and give back HTML/DOM for
preview and editing. Not implemented — the pipeline currently runs one way.

## Fidelity is best-effort

The mapping is **heuristic, not deterministic**. HTML/CSS and PPTX are different
formats with different primitives, and neither is a superset of the other: box
layout, text flow, filters, and blend modes have no exact OOXML equivalent, and
PowerPoint's own model (placeholders, theme colors, freeform geometry) has no
exact CSS equivalent.

So the contract is a *close* deck, not a pixel-identical one. Where a faithful
mapping is impossible, the converter picks the closest editable construct and
records a `Warning` rather than dropping content. Callers that need pixel
fidelity over editability can fall back to `convertDeckRaster`, which rasterizes
each slide into a full-bleed picture.

## Scope

- **In scope:** a documented subset of HTML/CSS aimed at slide layouts —
  sectioned slides, common Tailwind-shaped utility styling, iconify icons,
  gradients, tables, lists. Extend it by adding fixtures.
- **Out of scope:** rendering arbitrary web pages. This is not a browser.
- **Environment:** **browser only.** It needs a real DOM — iframe,
  `getComputedStyle`, `getBoundingClientRect`, canvas, fonts — and is not
  Node-portable. Layout tests run in headless Chromium (Playwright), not
  jsdom/happy-dom.

## Public API

```ts
import { convertDeck, convertSlide } from 'dom2pptx'

await convertDeck(fullHtmlString, opts) // → ConvertResult
await convertSlide(headHTML, slideHTML, opts) // → { model, warnings }
```

See `src/index.ts` for `ConvertOptions` — including the injectable `resolveIcon`
and `pptxFactory` seams, the `output` delivery mode, and the opt-in
`vectorizeSvg` — and `src/ir/model.ts` for the IR types.

## Development

```bash
pnpm install
pnpm run build        # tsdown → ESM dist/
pnpm run typecheck
pnpm run lint
pnpm run test         # builds, then runs unit tests
```

The writer dependency, `@shbernal/ts-pptx`, is consumed from public npm — no
local link or sibling checkout is required. This package is not itself published
during the prototype phase; consumers link it locally.
