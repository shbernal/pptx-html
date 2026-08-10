---
doc-schema-version: 1
title: 'Architecture'
summary: 'The two lanes, the two models, and the rules each part of the loop is built on.'
read_when:
  - Changing anything in src/import, src/render, src/parse or src/emit
  - Deciding where a new capability belongs
  - Wondering why two types that look alike are deliberately kept apart
doc_type: 'architecture'
---

# Architecture

One-way dependency: **consumer app → pptx-html → ts-pptx**. No cycles.

`pptx-html` does not know about AI or UI. `@shbernal/ts-pptx` is its only runtime
dependency, and consumers never import `ts-pptx` directly.

## Two lanes, and the directory layout says which is which

**The loop** — `src/import/` → `src/render/` → `src/parse/` → `src/emit/`, joined
by `src/loop.ts`, over the model in `src/ir/`. All of it is isomorphic: it drives
the read model's typed object graph and runs in Node and in Chromium alike. This
is the lane [Invariant R](./round-trip.md) is about.

**The heuristic lane** — `src/heuristic/`, entered by `convertDeck` /
`convertSlide`. DOM → an inferred model → the writer. **Browser-only** (iframe,
`getComputedStyle`, `getBoundingClientRect`, canvas, fonts), and best-effort by
construction.

They share the writer and nothing else. Do not let a code path serve both: the
two have different contracts, and merging them weakens the strong one without any
test going red.

`src/heuristic/model.ts` is **not** `RenderIr` and must not become it. Deriving a
paint model from a DOM would mean inventing node identity and placeholder
inheritance that the page does not have — inference dressed as fidelity — and
would put one type behind two incompatible guarantees. Three of its type names
(`Rect`, `TableCell`, `Background`) mean something different from the loop's,
which is why `src/index.ts` exports it under a `heuristic` namespace.

`src/heuristic/extractor.ts` holds the package's only `@ts-nocheck`. It is
stringified and `eval`'d inside the slide iframe, so it cannot be checked where it
runs; `src/heuristic/read.ts` validates its whole result against the model before
anything reaches the writer. That validation is the *entire* argument for
tolerating the suppression — adding a second unchecked consumer of the extractor's
output breaks it, and retyping the extractor in place would not strengthen it.

## Modeled, carried, or warned — never approximated

HTML/CSS and PPTX are different formats and neither is a superset of the other,
but that does not license guessing. Every element lands in exactly one of three
states:

- **Modeled** — the IR represents it; it survives the loop exactly.
- **Carried** — the IR does not model it, so the original XML moves across
  untouched. No approximation, no loss. This is the residual channel that lets
  best-effort coverage coexist with a hard round-trip guarantee.
- **Warned** — neither modeled nor carried, and the conversion says so. A visible
  failure beats a silent one.

The fourth state — **approximated**, output that looks about right but has no way
back — is what the charter rules out.

The machine-readable classification is upstream's `FidelityNote`
(`Disposition` × `Cause`) from `@shbernal/ts-pptx/script`, checked against
`knownNoteConstructs()`. **Do not coin a local `modeled`/`carried`/`unsupported`
enum** alongside it: two vocabularies for one concept is how the differ and the
renderer drift apart. The trio above is the plain-language framing of that same
classification.

`Cause` also triages where a fix belongs: `unread` / `unwritable` are upstream
bugs to file against `ts-pptx`; `unsupported` is a property of OOXML and will not
be fixed by more converter work here.

In the inference lane, still aim for the closest **editable** PowerPoint
construct, not a pixel match, and record a `Warning` rather than dropping content.

## Two IRs, on purpose

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

### Rules `RenderIr` is built on

Each of these will break something if ignored.

- **EMU is the stored unit**, integers, for every position and size; `Pt` in a
  field name means OOXML states it in points. Inches appear only at the edges via
  `inchesOf`/`emuOf`. Storing inches — or both — puts float noise on every
  element, which then reads as a difference on every diff.
- **JSON is the wire format.** Part of the model is embedded in the rendered
  document and parsed back, so: no `undefined` (an absent field is a missing key,
  the single spelling of "absent"), no `Date`/`Map`/`Set`, and **no `Uint8Array`**
  — media lives behind an `AssetRef` and is addressed by content hash, never
  re-embedded. `test/unit/ir-render.test.ts` walks the model and fails on any of
  these.
- **Node identity is derived from the source, never generated** —
  `s{slide}.sp{cNvPr@id}`, structural rather than hashed, so a second import of
  the same deck agrees with the first and the differ aligns sides by id instead of
  by position.
- **Absent means inherited, not default.** Every `RunProperties` field, plus a
  paragraph's `align`/`bullet`, and the `inherit` arm of `Fill` and `Stroke`. This
  is why `TextRun` carries **two** property sets: `props` is what the run itself
  stated (what emit reads), `resolved` is what to paint after the placeholder →
  layout → master chain is walked. Writing the resolved value into `props` renders
  identically and bakes a layout's 44pt title into every slide — the flattening
  trap in miniature.
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

## Import is mapping, not parsing

`src/import/` turns a `.pptx` into both models. It is **browser-capable** — the
same code runs in Chromium — and it holds to three rules:

- **No XML.** Everything comes through `@shbernal/ts-pptx/read`'s typed object
  graph. When the read model exposes no accessor, that is an upstream ask, not a
  licence to reach into `Shape.element_`. **No code in `src/` touches OOXML
  directly** — see [the decisions record](./decisions.md) for why the one place
  that did is gone.
- **The contract side is not reimplemented.** `readModelToIr` is called for
  `DeckIr`; only the render side is local.
- **Media identity is joined by content hash.** Upstream's asset names
  (`image1.png`) and the package's partnames (`/ppt/media/image-1-1.png`) have no
  published map between them, and two names for one image is how a picture stops
  being traceable through the loop. Hashing both sides is the only join that does
  not depend on an internal convention holding still.

**Opaque is not carried.** `OpaqueNode` / `render: 'placeholder'` means *this IR
cannot paint it*; `RenderSlide.source: 'carried'` means *the write API cannot
author it*. A plain chart is the case that separates them — it round-trips through
`addChart` perfectly and simply cannot be drawn in a browser.

### The template's furniture is `chrome`, and it is not `nodes`

A slide's own shape tree is `slide.shapes`. The bands, rules, logos and background
pictures that carry a deck's visual identity are not in it — they live on the
layout and the master, and in a PowerPoint-authored deck they are usually most of
what a reader recognises. `RenderSlide.chrome` holds them, mapped by the same
mappers as `nodes` (upstream returns one `AnyShape` union from all three trees) and
painted beneath them, master tier first.

Keeping them out of `nodes` is the load-bearing part, and there are four separate
reasons, any one of which is sufficient:

- `project` would offer them as editable, and an edit to a layout shape has
  nowhere to go — emit **binds to** the layout part rather than redrawing it.
- Node ids are scoped to one shape tree, so a layout shape and a slide shape can
  both be `p:cNvPr/@id` 2. `chromeNodeId` gives each tier its own namespace
  (`s1.layout.sp2`) precisely so the two can never be confused for one another.
- Deleting a node is in surface. Deleting a master shape is not expressible.
- `modelHash` would move for a reason that has nothing to do with the slide.

They are excluded from `project` and included in `freeze` — drawn, never
addressed. The renderer enforces the second half in the document as well as in the
model: chrome carries no `data-pxh-node`, and its runs carry no `data-pxh-run` and
no `contenteditable`.

Whether a slide gets any is `p:sld/@showMasterSp` — PowerPoint's *Hide background
graphics*. It comes from `AG_ChildSlide`, the attribute group each child tier
carries about the tier above it, so the slide's flag governs the layout's shapes
and the layout's own flag governs the master's. Both tiers' **placeholders** are
excluded regardless: a layout placeholder is a prompt for the slide's content, and
the slide's own shape already carries the geometry it inherited from it.

Upstream also reads these shapes, for a different output, and the two do not
overlap. `printStandaloneScript` re-authors a layout's furniture into
`defineSlideMaster({ objects })`, because a standalone script has no template to
bind to and would otherwise emit a deck wearing the wrong suit. This project
prints the **template-anchored** tier, where the layout part is already in the
package — so chrome here is read to be *drawn* and nothing else, and that tier
suppresses the `layout.`-prefixed notes upstream files about the rebuild.

## Two ways to carry a slide, and they are not interchangeable

`importSlide(source: Presentation, index)` needs the **live source package** and
reproduces the slide's rel graph intact.

`appendSlides` takes a `SlideSource` — a one-method interface over plain
serializable `ExtractedSlide`s — so it can be fed from bytes, but it costs
placeholder inheritance and re-resolves `schemeClr` against the destination theme.
That is free for v1's input domain (already concrete, absolutely positioned) and
fatal for the deferred real-deck tier.

`appendSlides` can position (`at?`); `importSlide` cannot. Ordering otherwise
follows call order.

## The editable surface

`src/ir/surface.ts` defines exactly what a human may change in the rendered HTML
and have honoured on the way back: **run text,
`bold`/`italic`/`underline`/`strike`/`sizePt`/`color`, a paragraph's `align`, and
deleting a node.** Everything else — moving a box, changing geometry, restyling a
table, reordering or inserting slides — is **detected as drift**, never
interpreted.

The test a property has to pass is a 1:1 write-API option, not usefulness.
`underline` and `strike` pass because `RunProperties` already models each as the
three-value subset the writer expresses, and every other
`ST_TextUnderlineType`/`ST_TextStrikeType` token is imported as a fidelity note
rather than rounded into one of the three. `fontFace` fails: a face that may not
exist on the target machine needs interpretation to get back into the deck.

Each of the three values matters, including the one that draws nothing. A run
that inherits an underline from its list style and states `u="none"` is *not*
underlined, so "explicitly off" and "says nothing" are different facts and the
surface keeps them apart — absence is how this model spells inherited, at every
level. In the panel that is a drop-down with four options rather than a
checkbox, because a checkbox has no way to say the difference.

Keeping them apart cost an upstream fix.
[ts-pptx#14](https://github.com/shbernal/ts-pptx/issues/14): `readModelToIr`
mapped `u="none"` and `strike="noStrike"` to `undefined`, so a deck that
*already* stated the explicit off lost it on the read leg — and declared no note
for it, which meant `diffDeckIr` compared two contract models both missing the
field and reported clean. Edits made *through* the surface were never affected,
because `parse/edits.ts` writes the option onto the `CallIr` directly; it was the
untouched run that lost it. The oracle could not see this by construction, for
the same reason it could not see ts-pptx#13, so it was pinned by a test
asserting the loss until the fix landed in `b16fb74b` — which the test then
caught, and now asserts survival instead. The contract model carries the tokens,
so the lanes can see a regression here on their own.

### The paragraph tier, and the one property in it

`align` is the whole of it, and `bullet` is the instructive absence. Both live on
`ParagraphProperties`, both look equally editable, and only one passes the 1:1
test. `align`'s four values are the write option's own, and omitting the option
writes no `a:pPr/@algn` — so *inherited*, *left* and *centre* are three states the
option can say. `bullet` models the same three-way distinction and the write API
can express two of them: an omitted `bullet` and `bullet: false` both emit an
explicit `<a:buNone/>`, so there is no way to author a paragraph that states
nothing about its bullet. Upstream declares the loss
(`text.bullet.inherited`), and the ask is
[ts-pptx#15](https://github.com/shbernal/ts-pptx/issues/15) — the same shape as
ts-pptx#10, which added a spelling for an inherited *fill*.

A control with an "inherited" position that silently wrote the explicit off would
be the worst kind of wrong: a suppressed bullet and an inherited-none paint
identically, so nothing would look broken until someone edited the master and the
slide stopped following it. So the property waits.

The tier also changed where a value is *written*. Paragraph properties ride on
runs in the write contract — `readModelToIr` replicates each paragraph's options
onto every one of its runs — and the writer groups that flat list back into
paragraphs, **starting a new one wherever two adjacent runs disagree about
`align`**. Setting the value on the first run of a three-run paragraph therefore
does not restyle that paragraph; it splits it in two, and the run count is
unchanged, so a count-based guard cannot see it. `parse/edits.ts` writes a
paragraph property onto every run of the paragraph for that reason, and
`test/oracle/loop.test.ts` asserts the paragraph count rather than the value alone.

It is a data structure, not prose, because two consumers read it: the renderer
makes those regions editable, and the return path decides what counts as drift.
Two hand-maintained copies of "what is editable" diverge silently, and the failure
is invisible in both directions.

- `project(ir)` is the editable view — the part that may change, addressed by node
  id (positional addressing misaligns the moment a user deletes a node, which the
  surface allows).
- `freeze(ir)` is its complement — the part that may not. Both are derived from
  the same list, so they stay complementary by construction. The test that matters
  is that an in-surface edit leaves `freeze` bit-identical while an out-of-surface
  one does not; keep it when touching either function.
- A property is in surface only if it maps 1:1 onto a write-API option. Anything
  needing interpretation to get back into the deck stays out — the return path
  must never guess.
- `RenderSlide.chrome` is outside it entirely — the layout and master shapes a
  slide inherits are drawn and frozen, never projected. See *The template's
  furniture is `chrome`* above.

## The return path

The rendered document has **two channels**, and only one of them is trusted.

- The **visual channel** (SVG and HTML) is allowed to approximate. A preset
  geometry with no local formula is drawn as a plain box, marked
  `data-pxh-approx`, and moved past. This costs nothing, and the reason is
  structural: the picture is drawn *from* the model rather than being it.
- The **JSON island** is not allowed to approximate. `src/parse/` reads that and
  the declared surface attributes, and nothing else. It never consults
  `getComputedStyle` — computed style is lossy (font fallback, sRGB normalization,
  sub-pixel rounding) and has no representation for placeholder inheritance,
  colour transforms, autofit mode or geometry adjust values.

### A few `RenderIr` fields are paint data and can never reach a deck

Most of the model exists so the emit path can rebuild the deck. A small, marked
minority exists only so the picture is right, and the two are worth telling
apart when reading the IR.

`TextBody.autofitFontScalePct` and `autofitLineSpaceReductionPct` are the current
example. They come off the read model, they decide the size text is painted at,
and they do not travel onward from here: emit folds edits into a `DeckIr` and
never reads `RenderIr`, nothing in the editable surface can change them, and
re-deriving one would mean measuring text.

They are worth reading as the category's cleanest case *because* the reason
narrowed. It used to be that the numbers were gone from the contract model
altogether — `readModelToIr` flattened a baked `<a:normAutofit fontScale="…"/>`
to `fit: 'shrink'`, so a deck came back painting text it had shrunk to 40% at
full size. That was [ts-pptx#13](https://github.com/shbernal/ts-pptx/issues/13),
and it is fixed: the `DeckIr` carries `fit: { type: 'shrink', fontScale,
lnSpcReduction }` now, and the round trip preserves it. So these fields are no
longer the only copy of anything — the deck keeps its own, by its own route. What
makes them paint data is the architecture, not a gap, and that is the durable
form of the category. Nothing may treat such a field as part of the round-trip
guarantee, and each one says so in its own doc comment.

`RenderSlide.chrome` is the other, and it is paint data for a different reason:
not because the contract model dropped it, but because the contract model never
needed it. The layout and master parts ride along in the template package that
emit binds against, so the furniture is reproduced by the binding rather than by
anything this project writes. Reading it changes the picture and can never change
the deck.

### Two hashes, because one would collapse two different events

- `modelHash` is over the island block's text **exactly as embedded**, escapes and
  all. A mismatch means the model was tampered with, and it **throws** — it is not
  a lane. Hashing the escaped text is forced, not stylistic: `escapeForScript` has
  no safe inverse, so a reader that un-escaped before hashing would corrupt
  exactly the models most likely to be adversarial.
- `surfaceHash` is over `project(ir)`. A mismatch means a sanctioned edit. Hash the
  projection and never the raw DOM, or every browser normalization — attribute
  order, whitespace, colour serialization — reads as an edit.

### Four lanes, decided per slide and always reported

`exact`, `reconciled` (a sanctioned edit, re-modeled), `drifted` (an edit outside
the surface — island value kept, edit discarded, warned), and `heuristic` (no
island at all; that is `convertDeck`, not a fallback inside `parseDeck`). A caller
must be able to see "slide 4 fell to heuristic" without opening the deck.

The `drifted` lane is scoped to what the surface reader can see — a removed run,
an unknown address, a duplicate, a malformed props attribute — and **not** to a
moved box or a recoloured path. That gap is not a hole: an out-of-surface DOM edit
is **inert, not dangerous**, because emit reads the island and the surface reading
and nothing else. A shape dragged in dev tools is not misapplied, it is not
applied. The loss is the user's edit, never the deck's fidelity. Closing the gap
would mean re-deriving the model from the DOM to compare against — putting
inference back into the trusted path to detect something that cannot hurt
anything.

### Emit takes the source package

`emitDeck(parsed, { source })`. Masters, layouts, theme and any carried slide's XML
live there and nowhere else, and embedding a whole `.pptx` in the HTML would make
every document at least as large as the deck it shows. The document carries the
edits; the caller supplies the substance — the same split as `AssetMode` one level
down. The two are proved to be the same deck before anything is applied: a fresh
import of the package must produce the document's `modelHash`, because an edit
addressed by node id means nothing against another deck's shape tree.

**Media rides on a render option, not in the model.** `renderDeck(ir, { assets })`
takes `'inline'` (default) or `'ref'`; the island itself carries `AssetRef` keys
and a hashed manifest and **never bytes**, so `modelHash` is mode-independent and
the visual channel hydrates from the asset block rather than storing a second
copy. Every resolved byte range is verified against the manifest's `sha256` before
use — an unverifiable asset throws; an unresolvable one is a warning.

## Scope

- **In scope:** a documented subset of HTML/CSS aimed at slide layouts (sectioned
  slides, utility-class styling, iconify, gradients, tables, lists). Expand it by
  adding fixtures, not by chasing arbitrary CSS.
- **Out of scope:** rendering arbitrary web pages. This is not a browser.
