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
- The IR (`src/ir/`) is the boundary between the internal layers:
  - `src/import/` (`.pptx` → IR) is **isomorphic** — it drives the read model's
    typed object graph and runs in Node and in Chromium alike.
  - `src/extract/` (DOM → IR) is **browser-only** (iframe, `getComputedStyle`,
    `getBoundingClientRect`, canvas, fonts).
  - `src/emit/` (IR → ts-pptx) is **pure and isomorphic** — unit-testable
    without a browser. This is where custGeom correctness lives.
- Keep the IR as the single source of truth for the slide model. Changing the IR
  shape touches both layers; do it deliberately.

## Two IRs, On Purpose

`src/ir/render.ts` (`RenderIr`) is the **paint** model. `DeckIr` from
`@shbernal/ts-pptx/script` is the **contract** model — it is what emit writes and
what `diffDeckIr` judges. Both are built from **one loaded `Presentation`**
(`src/import/deck.ts`), so they cannot disagree about the source deck.

They stay separate because `DeckIr.slides[].calls[].args` is `IrValue` — untyped
write-API option bags — and nothing can be drawn from a bag. Merging them would
either drag geometry into upstream's write-call vocabulary or drag write-API
arguments into the renderer, and the second is how a renderer starts silently
deciding what gets emitted. **Nothing in `src/emit/` may import `ir/render`**; the
moment something is emitted from `RenderIr`, there are two writers and the oracle
is judging one of them.

Rules `RenderIr` is built on, each of which something will break if ignored:

- **EMU is the stored unit**, integers, for every position and size; `Pt` in a
  field name means OOXML states it in points. Inches appear only at the edges via
  `inchesOf`/`emuOf`. Storing inches — or both — puts float noise on every
  element, which then reads as a difference on every diff.
- **JSON is the wire format.** Part of the model is embedded in the rendered
  document and parsed back, so: no `undefined` (an absent field is a missing key,
  the single spelling of "absent"), no `Date`/`Map`/`Set`, and **no
  `Uint8Array`** — media lives behind an `AssetRef` and is addressed by content
  hash, never re-embedded. `test/unit/ir-render.test.ts` walks the model and
  fails on any of these.
- **Node identity is derived from the source, never generated** — `s{slide}.sp{cNvPr@id}`,
  structural rather than hashed, so a second import of the same deck agrees with
  the first and the differ aligns sides by id instead of by position.
- **Absent means inherited, not default.** Every `RunProperties` field, plus a
  paragraph's `align`/`bullet`, and the `inherit` arm of `Fill` and `Stroke`.
  This is why `TextRun` carries **two** property sets: `props` is what the run
  itself stated (what emit reads), `resolved` is what to paint after the
  placeholder → layout → master chain is walked. Writing the resolved value into
  `props` renders identically and bakes a layout's 44pt title into every slide —
  the flattening trap in miniature.
- **`inherit` and `none` are different decks.** A shape with `a:noFill` is
  deliberately transparent; a shape that states no fill takes one from its style
  reference. Collapsing them loses a real distinction. (Import can currently only
  produce `inherit` for a *fill* — see the upstream asks — and both for a *line*.)
- **Placement is slide-absolute at every depth**, composed by the read model's
  `absoluteFrame`. A group child's `a:xfrm` is stated in its group's child space
  and is not directly placeable; doing that arithmetic per consumer is how every
  shape in a nested group ends up subtly displaced. `GroupNode` therefore carries
  no child-space transform of its own.
- **`render: 'drawn' | 'placeholder'` is a rendering fact, not a fidelity one.**
  Fidelity is `FidelityNote` and nothing else.
- **Do not model what cannot round-trip yet.** A half-drawn construct with no
  matching `FidelityNote` is worse than an honest carried one — the note is what
  makes a difference *declared* rather than a defect.

## Import Is Mapping, Not Parsing

`src/import/` turns a `.pptx` into both models. It is **browser-capable** — the
same code runs in Chromium — and it holds to three rules:

- **No XML.** Everything comes through `@shbernal/ts-pptx/read`'s typed object
  graph. When the read model exposes no accessor, that is an upstream ask, not a
  licence to reach into `Shape.element_`. Raw OOXML lives in `src/repair/` and
  nowhere else.
- **The contract side is not reimplemented.** `readModelToIr` is called for
  `DeckIr`; only the render side is local.
- **Media identity is joined by content hash.** Upstream's asset names
  (`image1.png`) and the package's partnames (`/ppt/media/image-1-1.png`) have no
  published map between them, and two names for one image is how a picture stops
  being traceable through the loop. Hashing both sides is the only join that does
  not depend on an internal convention holding still.

**Opaque is not carried.** `OpaqueNode` / `render: 'placeholder'` means *this IR
cannot paint it*; `RenderSlide.source: 'carried'` means *the write API cannot
author it*. A plain chart is the case that separates them — it round-trips
through `addChart` perfectly and simply cannot be drawn in a browser.

The `import → emit` lane (`test/oracle/script-lane.ts`) replays `DeckIr` through
the write API with no HTML involved. Both legs are upstream code, so a failure
there is an issue to file rather than a local fix — and a defect it catches can
never be misattributed to the renderer.

## The Editable Surface

`src/ir/surface.ts` defines exactly what a human may change in the rendered HTML
and have honoured on the way back: **run text, `bold`/`italic`/`sizePt`/`color`,
and deleting a node.** Everything else — moving a box, changing geometry,
restyling a table, reordering or inserting slides — is **detected as drift**,
never interpreted.

It is a data structure, not prose, because two consumers read it: the renderer
makes those regions editable, and the return path decides what counts as drift.
Two hand-maintained copies of "what is editable" diverge silently, and the
failure is invisible in both directions.

- `project(ir)` is the editable view — the part that may change, addressed by
  node id (positional addressing misaligns the moment a user deletes a node,
  which the surface allows).
- `freeze(ir)` is its complement — the part that may not. Both are derived from
  the same list, so they stay complementary by construction. The test that
  matters is that an in-surface edit leaves `freeze` bit-identical while an
  out-of-surface one does not; keep it when touching either function.
- A property is in surface only if it maps 1:1 onto a write-API option. Anything
  needing interpretation to get back into the deck stays out — the return path
  must never guess.

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
- **Oracle (`test/oracle/`, no DOM):** the round-trip gate. See below.
- **Browser/e2e (`test/browser/`, Playwright + headless Chromium):** load HTML
  slide fixtures and run `convertSlide`/`convertDeck` with `output:'base64'` and a
  static injected `resolveIcon`. jsdom/happy-dom are insufficient.
- Keep most assertions in the headless unit layer; minimize the e2e surface.

## The Round-Trip Oracle

`test/oracle/roundtrip.ts` decides whether a deck survived the loop, and it is
what makes Invariant R a claim rather than a hope. **A fidelity change is not
done until the oracle covers it.**

- **Equality is upstream's, not ours.** `readModelToIr` → `canonicalDeckIr` →
  `diffDeckIr`. Do not write a normalizer or a differ; zip order, timestamps,
  `rId`s and `cNvPr` ids are canonicalized away by design.
- **`undeclared` is the only field a gate may read.** `declared` is the fidelity
  contract working, `added` is the write path being explicit where the source was
  implicit, and `unmatchedNotes` are usually constructs the read model cannot see
  at all — report them, never gate on them, never "fix" one by deleting a note.
- **A lane declares its own losses.** `diffDeckIr`'s `notes` argument must be the
  notes of the tier that produced the output. `DeckIr.fidelity` is the wrong set
  (it describes the source, not the output), and `printScript`'s notes are right
  only for the printed-script tier — `scriptTierNotes()` exists so that case has a
  name and nothing else reaches for it by accident. A lane returning no notes is
  claiming to lose nothing, and is held to it.
- **Corpus (`test/corpus/decks.ts`)** is generated by the writer, never committed
  as binaries, and must stay deterministic — no clock, no filesystem, no
  randomness. Every deck is generated by `@shbernal/ts-pptx` itself because that
  *is* the v1 input domain; decks authored in PowerPoint are a deferred second
  tier. Add a deck when you add a construct.
- **The coverage snapshot is the progress metric.** Notes counted by
  `Disposition` × `Cause`, plus authored-vs-carried slides. It is expected to
  change; it must never change silently, which is why it is snapshotted.
- A green run on an empty deck proves nothing, so `assertNonTrivial` guards every
  corpus entry, and the harness is mutation-tested against itself in
  `roundtrip.test.ts`. Keep both when touching the harness.

## Verification

- For source changes, run `pnpm run build` and `pnpm run typecheck`.
- For behavior changes, run `pnpm run test:unit`.
- For anything touching fidelity, `pnpm run test:oracle` — and expect to update
  the coverage snapshot deliberately, never with a blind `-u`.
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
