# Changelog

Notable changes per release. Format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/); versions follow
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

Pre-1.0, the two lanes carry different promises and the version reflects that:
the loop's four legs (`importDeck` / `renderDeck` / `parseDeck` / `emitDeck`) and
the editable surface are the stable surface, and the heuristic lane's inferred
model is expected to move.

## [Unreleased]

### Fixed

- **A table cell's two diagonals are modeled and drawn.** `a:lnTlToBr` and
  `a:lnBlToTr` — PowerPoint's Diagonal Down/Up Border, and how a cell is struck
  out — were dropped on import with no `FidelityNote`, which is neither modeled,
  carried nor warned. They were also not a loss that had to happen: the read model
  decodes both (`CellBorders.tlToBr` / `.blToTr`) and the write API takes them back
  (`TableCellProps.diagonal`), so both sides of `@shbernal/ts-pptx` could express a
  diagonal and only this pipeline could not. `TableCell.borders` now carries all
  six lines and the renderer draws them, the two diagonals last, across the whole
  spanned region of a merged cell.

- **The document's language reaches the emitted deck.** The heuristic lane parses
  `<html lang>` and assigned it to `pptx.lang` and to `theme.lang`. The writer has
  neither: both landed as own properties on the instance and nothing wrote them
  into the package, so every deck came out `en-US` whatever the page said. `lang`
  is a *run* option, so it now travels on each `addText` call and on each table
  cell — a table-level one does not reach the cells, which is checked against the
  emitted XML rather than read off the type.

- **A gradient stop and a table cell's edge keep their colour's transform list.**
  Both arrived pre-flattened — a scheme token and a painted hex with nothing
  between them — so `transforms: []` meant "the reader had nothing to say", which
  is the same shape as "the deck stated none". A `lumMod`-darkened stop was
  indistinguishable from one that was not, and re-authoring either against a
  different theme silently stopped tracking it. The cell edge lost more than the
  list: its colour carried no `alpha`, so a transparent rule was imported opaque
  and drawn as one. Both now decode through the same `colorOf` a solid fill uses,
  on `ResolvedColor`s that arrived in `@shbernal/ts-pptx` 3.6.0
  ([ts-pptx#26](https://github.com/shbernal/ts-pptx/issues/26)).

- **An inherited italic is painted.** `italic` was the one run property with no
  resolved counterpart, so a run taking its slant from a master's `a:defRPr i="1"`
  rendered upright. Paint-only — `props` is what emit reads, so the deck always
  round-tripped, which is exactly why nothing caught it. `Run.resolvedItalic`
  ([ts-pptx#27](https://github.com/shbernal/ts-pptx/issues/27)) closes it, and
  `ResolvedRunProperties` now carries `italic` beside `bold`.

### Changed

- **`IR_VERSION` is 8.** `TableCell.borders` gained the two diagonals as required
  fields, so a document rendered by an older build carries no key to read them
  from. The parser refuses the mismatch and says to re-render, which is the true
  answer; nothing migrates.

### Added

- **`cell-diagonal` joins the corpus** as a primitive-tier deck: three cells, one
  with a single diagonal, one with both, one with none. The distinction is only
  testable from a generated deck if the deck states it three ways — a lane that
  carried the first cell's rule onto every cell would agree with a corpus that
  used one.

- **`color-transform` joins the corpus** as a primitive-tier deck, holding both
  halves of the flattening above: a themed gradient stop with a transform beside
  one without, and a table whose rules state an opacity. `transparency` is the one
  colour transform the write API can author, which is what makes the distinction
  reachable from a generated deck at all. `layout-placeholder`'s master now states
  italic for body level 1, so the corpus contains a run whose slant is written down
  nowhere but the master.

- **`cloneIr(ir)`** — a deep copy of a `RenderIr`, exported beside `emuOf` and
  `inchesOf`. It exists because the two spellings the repo was using are not
  equivalent in general: `JSON.parse(JSON.stringify(...))` drops
  `undefined`-valued keys and cannot carry a `Uint8Array`, and they agree on this
  model only because of its JSON contract. A consumer holding an IR needs the
  same guarantee, and the one that keeps working if that contract ever moves is
  `structuredClone`.

- **`EMU_PER_POINT`** joins `EMU_PER_INCH` on the public surface. Both are the
  model's stated units, and a consumer reading a `widthPt` beside a `Box` needs
  the second one to make sense of the pair. It was already in the package, one
  directory down in the renderer.

- **`eachNode(nodes, visit)` and `eachTextBody(nodes, visit)`** — the two
  traversals of a node tree, exported beside the model they walk. `eachTextBody`
  is the answer to "where does text live in a node tree", stated once and
  exhaustive over `RenderNode['kind']`, so a new kind is a compile error rather
  than an omission wherever the question is asked. `eachNode` visits a group
  before its children and lets a consumer prune a subtree by returning `false`,
  which is what "this node has no box, so nothing inside it was drawn" needs.

## [0.1.2] — 2026-08-11

### Fixed

- **An SVG picture is drawn as its SVG, not as the raster fallback beside it.**
  A vector picture is two parts in OOXML: the art hangs off the `asvg:svgBlip`
  extension, and `a:blip/@r:embed` holds a fallback for readers that cannot draw
  vectors. The importer took the fallback — and `@shbernal/ts-pptx` writes that
  fallback as a *1×1 transparent PNG*, so every icon in a deck rendered as one
  invisible pixel. Nothing warned, because a picture that resolves to a real part
  looks resolved: the SVG parts arrived in the asset manifest and no node
  referenced them. Found by the first outside consumer, on a deck where 28 of 49
  slides lost every glyph.

  The same change picks up the **SVG-only** form, where `a:blip/@r:embed` is
  legitimately absent (PowerPoint's Insert → Icons, and a plain SVG insert).
  Those pictures previously took the `dropped` arm and rendered as nothing at
  all.

  The raster fallback is still carried in the manifest, so a consumer that cannot
  draw SVG still has something to fall back to.

### Added

- **`picture-svg` joins the corpus** as a primitive-tier deck. The gap above was
  not a subtle mapping error; it was a construct the corpus did not contain, and
  a ledger cannot report on a construct nobody generates. The fidelity ledger now
  carries upstream's `image.svg` note, which states that the vector part is
  carried and the source's own fallback is regenerated rather than kept.

## [0.1.1] — 2026-08-10

The library is unchanged from 0.1.0. What changed is how it reaches you.

### Changed

- **Releases are published from CI, through npm's trusted publishing.** The
  workflow mints a short-lived OIDC token that npm exchanges for a credential
  scoped to this package and that one workflow file — no long-lived token exists
  to be stolen — and the package is stamped with **provenance**, so the npm page
  links the tarball to the commit and the run that built it. 0.1.0 was published
  from a laptop and carries none; this is the first version you can verify that
  way.
- Publishing is triggered by publishing the GitHub release rather than by
  pushing the tag, keeping the deliberate half of a release a human act.

## [0.1.0] — 2026-08-10

First public release.

### Added

- **The loop.** `importDeck`, `renderDeck`, `parseDeck` and `emitDeck` — `.pptx`
  to a paint model, to HTML, back to the model, back to `.pptx`. Anything the
  model does not express is carried across as its original XML rather than
  approximated, and anything that can be neither modeled nor carried is warned
  about.
- **Invariant R and the oracle that gates it.** A generated corpus of 26 decks
  runs the full loop on every CI build, compared under a normalized read model
  rather than byte-for-byte. The per-construct fidelity ledger is snapshotted so
  coverage cannot move without a visible diff.
- **The JSON island.** `renderDeck` embeds the model beside the visible SVG and
  `parseDeck` reads *that*, never `getComputedStyle` — which is what lets the
  visible channel approximate freely without any of it reaching the emitted deck.
  A document whose integrity hashes do not match is refused, not downgraded.
- **A declared editable surface.** `project` names what a rendered document may
  change and `freeze` is its complement; both are exported, so the question "what
  may I safely edit in this HTML?" is answerable without reading the renderer.
  It is run text, `bold` / `italic` / `underline` / `strike` / `sizePt` /
  `color`, a paragraph's `align` / `bullet` / `marginLeftPt` / `indentPt`, and
  deleting a node. A property is in only if a write-API option can author every
  one of its values, *including the absent one* — everything else is detected as
  drift rather than interpreted.
- **Lane reporting.** `parseDeck` states per slide how it read the document
  (`exact`, `reconciled`, `drifted`, `heuristic`) instead of leaving a caller to
  infer success from the fact that a file appeared.
- **A preview drawn from the model, never rasterized.** The renderer resolves
  twelve preset geometries exactly from ECMA-376's Annex D definitions and marks
  the rest as an obvious box rather than guessing an outline, and it draws the
  layout's and master's non-placeholder shapes as `chrome` — visible, and
  editable nowhere, because an edit to it would belong to every slide bound to
  that layout.
- **The heuristic lane.** `convertDeck` / `convertSlide` infer a model from a
  rendered DOM, for HTML this library did not produce. Browser-only,
  explicitly best-effort, and sharing no code path with the loop.
- **A public site**, at <https://shbernal.github.io/pptx-html/>, deployed from the
  same workflow that gates the library — the publish hangs off `needs: validate`,
  so a site whose round-trip oracle is failing cannot go out.
  - **A playground** that runs `importDeck → renderDeck → parseDeck → emitDeck` in
    the visitor's browser. The sample decks are the oracle's own corpus rather
    than decks written for the page, editing goes through `EDITABLE_SURFACE` and
    nothing else, the lane is reported per slide, and the round-tripped `.pptx`
    is downloadable beside the original. Nothing is uploaded anywhere.
  - **The fidelity ledger**, generated from the oracle's coverage reporter over
    the corpus, so the page cannot claim more than CI knows. Every row links to
    that deck running in the playground.
  - **`docs/` published as written.** `site/docs/` is a generated mirror; the
    tracked design record stays the source of truth, and the generator fails
    rather than warns when the two drift.
- **A runnable example.** `examples/round-trip.mjs` runs all four legs, edits a
  run, and re-imports the emitted deck to show the edit arrived. CI runs it,
  because it is the only thing here that imports the package by name and so the
  only check that the `exports` map works the way a consumer will use it.

### Notes on scope

- The guarantee covers decks written by `@shbernal/ts-pptx`, which is what the
  corpus is made of. Decks authored in PowerPoint are a deliberate second tier
  and are not yet gated.
- There is no rasterizer, on purpose. A flattened slide is the one output that
  can never re-enter the loop — see
  [docs/decisions.md](./docs/decisions.md).
- ESM only, Node `>=24`. `@shbernal/ts-pptx` `^3.2.0` is the one runtime
  dependency.

[unreleased]: https://github.com/shbernal/pptx-html/compare/v0.1.2...HEAD
[0.1.2]: https://github.com/shbernal/pptx-html/compare/v0.1.1...v0.1.2
[0.1.1]: https://github.com/shbernal/pptx-html/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/shbernal/pptx-html/releases/tag/v0.1.0
