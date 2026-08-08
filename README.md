# dom2pptx

Turn HTML from the DOM into PowerPoint (`.pptx`) decks, by way of
[`@shbernal/ts-pptx`](https://www.npmjs.com/package/@shbernal/ts-pptx).

This library owns the link between **HTML/DOM and `ts-pptx` calls**, in both
directions. It reads a document, builds an intermediate slide model (IR), and
drives the writer. It does not emit OOXML itself — that is `ts-pptx`'s job — and
it does not generate the HTML — that is the caller's job.

```
your app / AI agent               produces / displays HTML slides
        │  imports
        ▼
dom2pptx (this package)           HTML ⇄ IR ⇄ ts-pptx calls
        │  depends on
        ▼
@shbernal/ts-pptx                 IR-agnostic OOXML reader + emitter
```

## Why

The centre of the project is a **loop**, not a one-way conversion:

```
.pptx  ──import──►  IR  ──render──►  HTML   (what a human sees / edits)
  ▲                  ▲                 │
  └────emit──────────┴─────parse───────┘   (what a machine reads back)
```

**1. Preview and edit decks on the web.** HTML renders in a browser instantly;
PPTX does not. A deck goes out as HTML, gets looked at and edited there, and
comes back as a `.pptx` — the page is the preview *and* the editing surface.

**2. Let AI agents build decks.** Models are good at emitting HTML and bad at
emitting OOXML. Give an agent an HTML target and it can produce a deck; this
library is the adapter that makes the HTML land as editable slides.

The loop only means anything if it is lossless, which is the property the whole
design is arranged around:

> **Invariant R.** For any deck this pipeline can write, `import → render →
> parse → emit` produces a deck **equal under the normalized read model** to the
> input. Slides whose features the IR does not model are carried across
> **byte-identical** rather than approximated.

Equality is normalized, not byte-for-byte: zip entry order, timestamps,
relationship ids and element ids all vary legally, and the comparison
canonicalizes both sides before diffing.

**Status.** The forward half (HTML → IR → `.pptx`) is what ships today. The
import and return halves are under construction; until the round-trip oracle
gates CI, treat Invariant R as the charter, not as a shipped guarantee.

## Modeled, carried, or warned — never approximated

HTML/CSS and PPTX are different formats and neither is a superset of the other:
box layout, text flow, filters and blend modes have no exact OOXML equivalent,
and PowerPoint's own model (placeholders, theme colors, freeform geometry) has
no exact CSS equivalent. That does not make the output a guess. Every element
lands in exactly one of three states:

- **Modeled** — the IR represents it, and it survives the loop exactly.
- **Carried** — the IR does not model it, so its XML moves across untouched. No
  approximation, and no loss.
- **Warned** — it can be neither modeled nor carried, and the conversion says
  so. A visible failure, never a silent one.

What this rules out is the fourth state — *approximated*: content that comes out
looking about right but has no way back. That is why the `html2canvas` raster
fallback was removed rather than kept as an escape hatch. A slide flattened into
a picture is the one output that can never re-enter the loop, so producing a file
that way is a failure wearing a success's clothes.

Inference is still how the **secondary** lane works: HTML that carries no IR of
its own is read from the rendered DOM, and that reading is genuinely heuristic.
It stays honest by the same rule — closest editable construct, plus a `Warning`.

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
