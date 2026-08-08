# AGENTS.md

Guidance for coding agents working in the `dom2pptx/` project.

## Repository Expectations

- This repo builds `dom2pptx`: it moves slides between HTML and PPTX by driving
  `@shbernal/ts-pptx`. Its role is the **HTML ⇄ ts-pptx** link, both directions.
- This is a **standalone top-level project** (`~/dev/dom2pptx`) with its own git.
  It is not part of a workspace; run its commands from this directory.
- Use `pnpm`. Node `>=24`. Keep source in `src/`, tests in `test/`. Treat `dist/`
  as generated build output.
- Preserve unrelated dirty state. Do not revert user changes.

## Purpose

**Round-tripping is the centre of the project**, not a deferred extra:

```
.pptx  ──import──►  IR  ──render──►  HTML   (what a human sees / edits)
  ▲                  ▲                 │
  └────emit──────────┴─────parse───────┘   (what a machine reads back)
```

The property the design defends:

> **Invariant R.** For any deck this pipeline can write, `import → render →
> parse → emit` produces a deck **equal under the normalized read model** to the
> input. Slides whose features the IR does not model are carried across
> **byte-identical** rather than approximated.

Equality is normalized, not byte-level — zip order, timestamps, `rId`s and
`cNvPr` ids all vary legally, so both decks are canonicalized before diffing.
Byte determinism is explicitly *not* a goal and is not an upstream ask.

Two things this serves, in priority order:

1. **Web previews and web editing.** HTML is the medium a deck is looked at and
   edited in; this library is what makes that page redeemable as a real `.pptx`.
2. **AI-authored decks.** Agents emit HTML well and OOXML badly. This is the
   adapter that makes agent-emitted HTML land as editable slides.

The **heuristic DOM → IR lane stays**, but it is now the *secondary* path: it is
for HTML that carries no IR of its own. HTML produced by this library's own
renderer carries its IR in the document, and the return path parses that IR
rather than re-deriving it from `getComputedStyle` — computed style is lossy and
has no representation for placeholder inheritance, colour transforms, autofit
mode or geometry adjust values.

## Modeled, Carried, or Warned — Never Approximated

HTML/CSS and PPTX are different formats and neither is a superset of the other,
but that does not license guessing. Every element lands in exactly one of three
states, and "degrade and warn" is replaced by **carry, or warn**:

- **Modeled** — the IR represents it; it survives the loop exactly.
- **Carried** — the IR does not model it, so the original XML moves across
  untouched. No approximation, no loss. This is the residual channel that lets
  best-effort coverage coexist with a hard round-trip guarantee.
- **Warned** — neither modeled nor carried, and the conversion says so. A
  visible failure beats a silent one.

The fourth state — **approximated**, output that looks about right but has no
way back — is what the charter rules out.

- The machine-readable classification is upstream's `FidelityNote`
  (`Disposition` × `Cause`) from `@shbernal/ts-pptx/script`, checked against
  `knownNoteConstructs()`. **Do not coin a local `modeled`/`carried`/`unsupported`
  enum** alongside it: two vocabularies for one concept is how the differ and the
  renderer drift apart. The trio above is the charter's plain-language framing of
  that same classification.
- `Cause` also triages where a fix belongs: `unread` / `unwritable` are upstream
  bugs to file against `ts-pptx`; `unsupported` is a property of OOXML and will
  not be fixed by more converter work here.
- In the inference lane, still aim for the closest **editable** PowerPoint
  construct, not a pixel match, and record a `Warning` rather than dropping
  content.
- Do not add deterministic-looking guarantees to the docs beyond what the
  round-trip oracle actually gates. Fidelity claims must stay honest.

## The Raster Path Was Removed On Purpose

`convertDeckRaster` / `rasterizeSlide` and the `html2canvas` dependency are gone,
and **nothing should reintroduce a rasterizer in `src/`**. A slide flattened into
a full-bleed picture has no model to re-import — it is the one output that can
never re-enter the loop, so it is a failure wearing a success's clothes. The cost
is real and was accepted: hostile HTML no longer always yields *some* file. A
warned failure beats a file that can never come back.

The obvious follow-up — "then render `.pptx` → PNG directly" — was asked and
answered: it cannot run in a browser. It needs a PowerPoint-grade renderer, and
the slides that would most want a preview are the carried ones, which by
definition have no model to draw from. The real options (PowerPoint COM
`Slide.Export`, LibreOffice headless) are out-of-process and host-dependent, so a
preview generator is a **test utility, not a package feature**. One under `test/`
for visual review of the corpus is welcome; nothing goes in `src/`.

Canvas use that remains in `src/extract/` and `src/emit/` is *not* this: it
rasterizes individual CSS gradients and re-encodes images, and each of those is a
modeled IR element. It is not slide rasterization.

## Architecture & Boundaries

One-way dependency: **consumer app → dom2pptx → ts-pptx**. No cycles.

- `dom2pptx` does **not** know about AI or UI. `@shbernal/ts-pptx` is its only
  runtime dependency; consumers never import `ts-pptx` directly.
- The IR (`src/ir/model.ts`) is the boundary between the two internal layers:
  - `src/extract/` (DOM → IR) is **browser-only** (iframe, `getComputedStyle`,
    `getBoundingClientRect`, canvas, fonts).
  - `src/emit/` (IR → ts-pptx) is **pure and isomorphic** — unit-testable
    without a browser. This is where custGeom correctness lives.
- Keep the IR as the single source of truth for the slide model. Changing the IR
  shape touches both layers; do it deliberately.

## Scope

- **In scope:** a documented subset of HTML/CSS aimed at slide layouts (sectioned
  slides, utility-class styling, iconify, gradients, tables, lists). Expand it by
  adding fixtures, not by chasing arbitrary CSS.
- **Out of scope:** rendering arbitrary web pages. This is not a browser.

## Fix Upstream When Possible

- Prefer fixing generic OOXML / emitter problems upstream in `@shbernal/ts-pptx`
  (it helps every consumer) over patching them here. When a fix belongs upstream
  but is not yet released, keep any stopgap here thin and clearly marked, and
  drop it once a release carries the fix.
- Anything discovered while building the custGeom/SVG-path vectorizer (a missing
  custGeom case, a measure gap) goes upstream, not patched locally.

## Testing

- **Unit (`test/unit/`, no DOM, fast):** feed hand-written `SlideModel` fixtures
  into `emit/*`, write to base64, and parse back with ts-pptx's `read`/`inspect`
  to assert structure. Prefer structural assertions over binary goldens.
- **Browser/e2e (`test/browser/`, Playwright + headless Chromium):** load HTML
  slide fixtures and run `convertSlide`/`convertDeck` with `output:'base64'` and a
  static injected `resolveIcon`. jsdom/happy-dom are insufficient.
- Keep most assertions in the headless unit layer; minimize the e2e surface.

## Verification

- For source changes, run `pnpm run build` and `pnpm run typecheck`.
- For behavior changes, run `pnpm run test:unit`.
- Lint and format with `pnpm run lint` / `pnpm run format:check`.

## Reference Order

Most questions here are answered by ts-pptx's own API surface, not by the OOXML
spec — this package drives a writer, it does not emit XML. Check in this order:

1. **ts-pptx's shipped type declarations.** The package publishes its API as
   `.d.ts` files next to `dist/`. Locate it with
   `node -e "console.log(require.resolve('@shbernal/ts-pptx/package.json'))"`
   (pnpm hides the real directory under `node_modules/.pnpm/`), then read the
   `exports` map. The subpaths that matter here:
   - `@shbernal/ts-pptx` — the writer: `addShape` / `addText` / `addTable` /
     `addImage` option shapes, `ShapeType.custGeom`, and the freeform point DSL
     that `src/emit/custgeom.ts` passes through unchanged.
   - `@shbernal/ts-pptx/read` — `Presentation.load`, the round-trip oracle the
     unit tests already assert against.
   - `@shbernal/ts-pptx/inspect` — per-element view (box, fill, text runs,
     paragraph boundaries, `a:bodyPr` autofit mode). Use it to confirm what an
     emit path or a repair actually produced.
   - `@shbernal/ts-pptx/measure` — font metrics and measured text fit, for any
     mapping decision that depends on whether text will fit its frame.
   - `@shbernal/ts-pptx/zip` — the fflate ZIP toolkit `src/repair/` uses. Do not
     add a separate ZIP dependency.
2. **The `ooxml` MCP** (ECMA-376 schema/spec) — for raw XML only, see below.
3. **Web search** — last resort.

Do not vendor large spec text into the repo.

## Raw OOXML Work

`src/repair/repair.ts` is the only code here that touches OOXML directly,
rewriting `ppt/slides/slideN.xml`, `presentation.xml`, `[Content_Types].xml`,
and the `.rels` parts after the writer has run. Everything else goes through
ts-pptx's DSL and should stay that way.

- Consult the `ooxml` MCP before changing any of those transforms. Children of
  `a:bodyPr` are a schema-ordered sequence, attributes like `p:sldSz@type` are
  enumerations, and dropping relationship entries can leave a package
  PowerPoint refuses to open.
- Verify, do not eyeball: write a deck, run the repair, read it back with
  `read` / `inspect` in a unit test. The file is `@ts-nocheck` and the
  transforms are regex over XML, so nothing else will catch a mistake.
- `repairPptxBase64` catches every failure and returns the *unrepaired* deck, so
  a broken repair degrades silently. That is why the round-trip assertion is not
  optional.
- These repairs are stopgaps for ts-pptx bugs (autofit defaults, duplicate
  `cNvPr` ids, generated notes masters, missing slide-size type). Per **Fix
  Upstream When Possible**, prefer moving one upstream to extending it here, and
  delete the local repair once a release carries the fix.
