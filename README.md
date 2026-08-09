# pptx-html

**Edit a PowerPoint deck as a web page, and get the deck back — not an
approximation of it.**

`pptx-html` moves slides between HTML and PPTX by driving
[`@shbernal/ts-pptx`](https://www.npmjs.com/package/@shbernal/ts-pptx). It owns
the HTML ⇄ `ts-pptx` link in both directions: it reads a `.pptx` into a slide
model, renders that model as HTML, reads the edited HTML back, and writes a
`.pptx` out again.

**[Run the loop in your browser →](https://shbernal.github.io/pptx-html/playground)**
Pick a deck, edit it as a web page, download the `.pptx` that comes back, and open
both files. No server, no upload, and no picture of a slide anywhere on the page.

## Why this exists

Two problems, one shape.

**Decks are trapped in a desktop application.** HTML renders in a browser
instantly; PPTX does not. If you want to preview a deck on the web — or let
someone edit it there — you have to leave the format, and everything that
converts a deck to HTML converts it *away*: the result looks about right and
cannot become a deck again.

**AI agents write good HTML and bad OOXML.** Ask a model for a slide and it will
produce clean markup. Ask it for a `.pptx` and it will produce a corrupt zip.
The adapter that turns the first into the second is the missing piece.

Both want the same thing: a conversion that survives the trip home.

## The loop

```text
.pptx  ──import──►  IR  ──render──►  HTML   (what a human sees / edits)
  ▲                  ▲                 │
  └────emit──────────┴─────parse───────┘   (what a machine reads back)
```

The loop only means anything if it is lossless, and that is the property the
whole design is arranged around:

> **Invariant R.** For any deck this pipeline can write, `import → render → parse
> → emit` produces a deck **equal under the normalized read model** to the input.
> Slides whose features the IR does not model are carried across
> **byte-identical** rather than approximated.

Equality is normalized, not byte-for-byte: zip entry order, timestamps,
relationship ids and element ids all vary legally, and both sides are
canonicalized before diffing. This is not an aspiration in a design doc — a
generated corpus of 17 decks runs the full loop on every CI build, and the
per-construct fidelity ledger is snapshotted so it cannot move silently. That
ledger is [published](https://shbernal.github.io/pptx-html/docs/fidelity),
generated from the oracle's own reporter, with every row linking to the deck
running in the browser.

## Modeled, carried, or warned — never approximated

HTML/CSS and PPTX are different formats and neither is a superset of the other.
Box layout, text flow, filters and blend modes have no exact OOXML equivalent,
and PowerPoint's own model — placeholders, theme colours, freeform geometry — has
no exact CSS equivalent. That does not make the output a guess. Every element
lands in exactly one of three states:

- **Modeled** — the IR represents it, and it survives the loop exactly.
- **Carried** — the IR does not model it, so its XML moves across untouched. No
  approximation, and no loss.
- **Warned** — it can be neither modeled nor carried, and the conversion says so.
  A visible failure, never a silent one.

What this rules out is the fourth state — *approximated*: content that comes out
looking about right but has no way back. It is why the `html2canvas` raster
fallback was removed rather than kept as an escape hatch. A slide flattened into a
picture is the one output that can never re-enter the loop, so producing a file
that way is a failure wearing a success's clothes.

## Install

```bash
pnpm add pptx-html
```

The writer, [`@shbernal/ts-pptx`](https://www.npmjs.com/package/@shbernal/ts-pptx),
comes with it.

## What it looks like

Four legs, one loop:

```ts
import { importDeck, renderDeck, parseDeck, emitDeck } from 'pptx-html'

const { render } = await importDeck(pptxBytes)
const { html } = await renderDeck(render, { bytes: pptxBytes }) // editable HTML
const parsed = await parseDeck(html) // the model back, plus a lane per slide
const { bytes } = await emitDeck(parsed, { source: pptxBytes }) // → .pptx
```

`renderDeck` writes the model into the document as a JSON island beside the
visible SVG, and `parseDeck` reads *that* — never `getComputedStyle`. It reports
per slide which lane it took (`exact`, `reconciled`, `drifted`, `heuristic`) and
throws rather than guessing when a document's integrity hashes do not match.

What a rendered document may be edited in is **declared, not implied**: run text,
`bold` / `italic` / `sizePt` / `color`, and node deletion. `project(ir)` is that
surface and `freeze(ir)` is its complement; both are exported, so "what may I
safely edit in this HTML?" is answerable without reading the renderer.

For HTML this library did not render, there is a second, explicitly best-effort
lane that infers a model from the rendered DOM:

```ts
import { convertDeck, convertSlide } from 'pptx-html'

await convertDeck(fullHtmlString, opts) // → ConvertResult
await convertSlide(headHTML, slideHTML, opts) // → { model, warnings }
```

## Status

All four legs of the loop are implemented and exported, and the round-trip oracle
gates CI. What that guarantee currently covers:

- **Input domain.** Decks written by `@shbernal/ts-pptx` — that is what the
  generated corpus is made of. Decks authored in PowerPoint are a deliberate
  second tier and are not yet gated.
- **Environment.** The loop is host-agnostic and runs in Node and the browser
  alike. The heuristic lane is **browser only** — it needs a real DOM (iframe,
  `getComputedStyle`, `getBoundingClientRect`, canvas, fonts), and its tests run
  in headless Chromium, not jsdom.
- **Distribution.** Published to npm as `pptx-html`. ESM only, Node `>=24`.
  Pre-1.0: the loop's four legs are stable, the heuristic lane's model is not.

Fidelity claims in these docs are held to what the oracle actually gates. Where
something is not covered, it says so.

## Scope

- **In scope:** a documented subset of HTML/CSS aimed at slide layouts —
  sectioned slides, common utility-class styling, iconify icons, gradients,
  tables, lists. Extend it by adding fixtures.
- **Out of scope:** rendering arbitrary web pages. This is not a browser.

## Further reading

- [examples/round-trip.mjs](./examples/round-trip.mjs) — the loop end to end and
  runnable: `pnpm run build && pnpm run example`. It edits a run, writes the deck
  back out, and re-imports it to show the edit arrived.
- [docs/](./docs/index.md) — the design record: [Invariant R and the
  oracle](./docs/round-trip.md), [architecture](./docs/architecture.md), and
  [decisions that must not be undone](./docs/decisions.md). Published at
  <https://shbernal.github.io/pptx-html/docs/>, alongside the playground and the
  fidelity ledger.
- [CONTRIBUTING.md](./CONTRIBUTING.md) — install, build, the three test layers,
  and what to run for which kind of change.
- [CHANGELOG.md](./CHANGELOG.md) — what changed, per release.

## License

[MIT](./LICENSE) © shbernal
