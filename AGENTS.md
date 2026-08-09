# AGENTS.md

Guidance for coding agents working in the `dom2pptx/` project.

## Repository Expectations

- This repo builds `dom2pptx`: it moves slides between HTML and PPTX by driving
  `@shbernal/ts-pptx`. Its role is the **HTML ⇄ ts-pptx** link, both directions.
- This is a **standalone top-level project** (`~/dev/dom2pptx`) with its own git.
  It is not part of a workspace; run its commands from this directory.
- Use `pnpm`. Node `>=24`. Keep source in `src/`, tests in `test/`. Treat `dist/`
  as generated build output.
- **Skills live in `.agents/skills/`** — the tracked source of truth, runtime
  agnostic. `.claude/skills/` is a junction to it (`.gitignore`d), so Claude Code
  loads the same files every other runtime does and there is one copy to edit.
  Recreate it after a fresh clone:

  ```powershell
  New-Item -ItemType Junction -Path .claude\skills -Target .agents\skills
  ```

  Repo-scoped skills only — anything not specific to this project belongs in the
  personal skills repo instead.
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

Canvas use that remains in `src/heuristic/` is *not* this: it rasterizes
individual CSS gradients and re-encodes images, and each of those is a modeled
element. It is not slide rasterization.

## Architecture & Boundaries

One-way dependency: **consumer app → dom2pptx → ts-pptx**. No cycles.

- `dom2pptx` does **not** know about AI or UI. `@shbernal/ts-pptx` is its only
  runtime dependency; consumers never import `ts-pptx` directly.
- **`src/` is two lanes, and the directory layout says which is which.**
  - The **loop** — `src/import/` → `src/render/` → `src/parse/` → `src/emit/`,
    joined by `src/loop.ts`, over the model in `src/ir/`. All of it is
    isomorphic: it drives the read model's typed object graph and runs in Node
    and in Chromium alike. This is the lane Invariant R is about.
  - The **heuristic lane** — `src/heuristic/`, entered by `convertDeck` /
    `convertSlide`. DOM → an inferred model → the writer. **Browser-only**
    (iframe, `getComputedStyle`, `getBoundingClientRect`, canvas, fonts), and
    best-effort by construction.
  - They share the writer and nothing else. Do not let a code path serve both:
    the two have different contracts, and merging them weakens the strong one
    without any test going red.
- `src/heuristic/model.ts` is **not** `RenderIr` and must not become it. Deriving
  a paint model from a DOM would mean inventing node identity and placeholder
  inheritance that the page does not have — inference dressed as fidelity — and
  would put one type behind two incompatible guarantees. Three of its type names
  (`Rect`, `TableCell`, `Background`) mean something different from the loop's,
  which is why `src/index.ts` exports it under a `heuristic` namespace.
- `src/heuristic/extractor.ts` holds the package's only `@ts-nocheck`. It is
  stringified and `eval`'d inside the slide iframe, so it cannot be checked where
  it runs; `src/heuristic/read.ts` validates its whole result against the model
  before anything reaches the writer. That validation is the *entire* argument
  for tolerating the suppression — adding a second unchecked consumer of the
  extractor's output breaks it, retyping the extractor in place would not
  strengthen it.

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

`src/parse/edits.ts` is the **one** seam where the two models meet. It is
one-directional, it writes nothing that did not change, and it carries only the
delta. Two consequences worth protecting: the unedited loop exercises none of its
structural mapping, so Invariant R for that case reduces to `import → emit`
(already gated); and the two models stay free to disagree about defaults, which
they do — `readModelToIr` resolves a run's colour to `000000` where the paint
model correctly records that the run stated nothing. The `calls[i]` ↔ `nodes[i]`
alignment it depends on is **asserted** (length plus `sourceName`), never trusted:
both sides walk `slide.shapes` in document order, but that is a property of two
independent traversals rather than a contract either publishes, and patching the
run that happened to line up would put a user's text into another shape and look
exactly like success.

**Emit is a pure function of the IR.** Nothing in `src/emit/` or `src/loop.ts`
measures text, reads a font metric or consults the host — measurement is resolved
*into* a model where it is taken (`src/heuristic/`) and never recomputed at emit
time. Otherwise the same IR yields different decks on different machines and
Invariant R becomes machine-dependent. `test/oracle/determinism.test.ts` keeps a
cross-process snapshot of the whole loop precisely because "deterministic by
construction" is a claim, not a check.

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
  licence to reach into `Shape.element_`. **No code in `src/` touches OOXML
  directly** — see *Raw OOXML Work* below for why the one place that did is gone.
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

## The Return Path

The rendered document has **two channels**, and only one of them is trusted.

- The **visual channel** (SVG and HTML) is allowed to approximate. A preset
  geometry with no local formula is drawn as a plain box, marked
  `data-d2p-approx`, and moved past. This costs nothing, and the reason is
  structural: the picture is drawn *from* the model rather than being it.
- The **JSON island** is not allowed to approximate. `src/parse/` reads that and
  the declared surface attributes, and nothing else. It never consults
  `getComputedStyle` — computed style is lossy (font fallback, sRGB
  normalization, sub-pixel rounding) and has no representation for placeholder
  inheritance, colour transforms, autofit mode or geometry adjust values.

**Two hashes, because one would collapse two different events.**

- `modelHash` is over the island block's text **exactly as embedded**, escapes
  and all. A mismatch means the model was tampered with, and it **throws** — it
  is not a lane. Hashing the escaped text is forced, not stylistic:
  `escapeForScript` has no safe inverse, so a reader that un-escaped before
  hashing would corrupt exactly the models most likely to be adversarial.
- `surfaceHash` is over `project(ir)`. A mismatch means a sanctioned edit. Hash
  the projection and never the raw DOM, or every browser normalization —
  attribute order, whitespace, colour serialization — reads as an edit.

**Four lanes, decided per slide and always reported**: `exact`, `reconciled`
(a sanctioned edit, re-modeled), `drifted` (an edit outside the surface —
island value kept, edit discarded, warned), `heuristic` (no island at all;
that is `convertDeck`, not a fallback inside `parseDeck`). A caller must be able
to see "slide 4 fell to heuristic" without opening the deck.

The `drifted` lane is scoped to what the surface reader can see — a removed run,
an unknown address, a duplicate, a malformed props attribute — and **not** to a
moved box or a recoloured path. That gap is not a hole: an out-of-surface DOM
edit is **inert, not dangerous**, because emit reads the island and the surface
reading and nothing else. A shape dragged in dev tools is not misapplied, it is
not applied. The loss is the user's edit, never the deck's fidelity. Closing the
gap would mean re-deriving the model from the DOM to compare against — putting
inference back into the trusted path to detect something that cannot hurt
anything.

**Emit takes the source package** (`emitDeck(parsed, { source })`). Masters,
layouts, theme and any carried slide's XML live there and nowhere else, and
embedding a whole `.pptx` in the HTML would make every document at least as large
as the deck it shows. The document carries the edits; the caller supplies the
substance — the same split as `AssetMode` one level down. The two are proved to
be the same deck before anything is applied: a fresh import of the package must
produce the document's `modelHash`, because an edit addressed by node id means
nothing against another deck's shape tree.

**Media rides on a render option, not in the model.** `renderDeck(ir, {assets})`
takes `'inline'` (default) or `'ref'`; the island itself carries `AssetRef` keys
and a hashed manifest and **never bytes**, so `modelHash` is mode-independent and
the visual channel hydrates from the asset block rather than storing a second
copy. Every resolved byte range is verified against the manifest's `sha256`
before use — an unverifiable asset throws; an unresolvable one is a warning.

## Two Ways To Carry A Slide, And They Are Not Interchangeable

`importSlide(source: Presentation, index)` needs the **live source package** and
reproduces the slide's rel graph intact. `appendSlides` takes a `SlideSource` — a
one-method interface over plain serializable `ExtractedSlide`s — so it can be fed
from bytes, but it costs placeholder inheritance and re-resolves `schemeClr`
against the destination theme. That is free for v1's input domain (already
concrete, absolutely positioned) and fatal for the deferred real-deck tier.
`appendSlides` can position (`at?`); `importSlide` cannot. Ordering otherwise
follows call order.

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
- **File the issue in the same unit of work that found the gap** — on
  `shbernal/ts-pptx`, before the commit, not batched into a later sweep. A gap
  that lives only in a chat session or a doc does not exist: sessions end and
  scratch plans are deleted by design, and the tracker is the one queue that
  survives both. An ask that turns out to be unclear gets closed, which is cheap;
  a gap never filed costs a rediscovery, usually by whoever re-invents the same
  workaround. Every stopgap therefore carries its **issue URL** and the condition
  under which it is deleted — see `src/import/paint.ts` and
  `test/oracle/script-lane.ts` for the shape of that comment.
- `.agents/skills/ts-pptx-upstream/` is the normative reference: what to file,
  what not to (anything about *HTML* stays here), how to write an ask that is
  still actionable months later, and what to do when a release lands.
- Anything discovered while building the custGeom/SVG-path vectorizer (a missing
  custGeom case, a measure gap) goes upstream, not patched locally.

## Testing

- **Unit (`test/unit/`, no DOM, fast):** the pure pieces of both lanes — the paint
  model and its surface, the island, the reconcile fold and the extractor
  boundary; plus hand-written `heuristic/` fixtures written to base64 and parsed
  back with ts-pptx's `read`/`inspect`. Prefer structural assertions over binary
  goldens.
- **Oracle (`test/oracle/`, no DOM):** the round-trip gate. See below.
- **Browser/e2e (`test/browser/`, Playwright + headless Chromium):** the two
  things that cannot be faked — reading the editable surface back out of a
  rendered document, and the heuristic lane's extract path
  (`convertSlide`/`convertDeck` with `output:'base64'` and a static injected
  `resolveIcon`). jsdom/happy-dom are insufficient.
- Keep most assertions in the headless layers; minimize the e2e surface.
- **Assert on the lane a slide took, not just on the file that came out.** Three
  bugs in the return path each made the system *quietly do nothing* and each
  produced a plausible, working document: the parser handing back the island as
  the edited model (so every edit was invisible), a colour compared by reference
  (so every untouched slide read as edited), a colour validated as a string (so
  every colour-bearing run raised an anomaly). A lane that degrades gracefully
  hides its own bugs; the output alone will not tell you.

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
- For anything touching the renderer, the surface reading or the heuristic lane,
  `pnpm run test:browser`.
- Lint and format with `pnpm run check`.
- **After changing what `src/index.ts` exports, read `dist/index.d.ts`.** The
  public surface has a failure mode neither `tsc` nor the tests can see: nothing
  in-repo imports through the entry point, and `export type *` from a module that
  also exports values emits those values as *types named after functions* —
  declared, uncallable, and discovered only by a consumer.

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
     that `src/heuristic/custgeom.ts` passes through unchanged.
   - `@shbernal/ts-pptx/read` — `Presentation.load`, the round-trip oracle the
     unit tests already assert against.
   - `@shbernal/ts-pptx/inspect` — per-element view (box, fill, text runs,
     paragraph boundaries, `a:bodyPr` autofit mode). Use it to confirm what an
     emit path actually produced.
   - `@shbernal/ts-pptx/measure` — font metrics and measured text fit. Read-only,
     and **not from the emit path**: a measurement taken at emit time makes the
     same IR produce different decks on different machines. Measuring belongs in
     `src/heuristic/`, resolved into that lane's model where it is taken.
   - `@shbernal/ts-pptx/zip` — the fflate ZIP toolkit, if a test ever needs to
     look inside a package. Do not add a separate ZIP dependency.
2. **The `ooxml` MCP** (ECMA-376 schema/spec) — for raw XML only, see below.
3. **Web search** — last resort.

Do not vendor large spec text into the repo.

## Raw OOXML Work

**There is none, and there must not be another.** Everything goes through
ts-pptx's DSL. `src/repair/repair.ts` used to rewrite `ppt/slides/slideN.xml`,
`presentation.xml`, `[Content_Types].xml` and the `.rels` parts after the writer
had run; it was measured against ts-pptx 3.0.0 across the whole corpus and
deleted. Four of its five rules never fired, and both that did were wrong — one
flattened an inherited autofit into an explicit one, the other **deleted every
speaker note in the deck**. It also swallowed exceptions and returned the
unrepaired file, so its output depended on whether something threw.

The lesson is worth more than the code was: **a layer that silently "fixes"
things hides both its own obsolescence and its own damage.** Every one of those
defects was invisible from outside — the deck opened, the text was there, and
only the notes were gone.

`test/oracle/writer-output.test.ts` now asserts the conditions the module existed
for do not occur. If the writer regresses, the answer is an upstream issue and a
failing test, **not a second repair pass**. If you find yourself reaching for a
post-write rewrite, read that file first — it names each rule and why it went.
