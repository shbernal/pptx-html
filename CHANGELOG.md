# Changelog

Notable changes per release. Format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/); versions follow
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

Pre-1.0, the two lanes carry different promises and the version reflects that:
the loop's four legs (`importDeck` / `renderDeck` / `parseDeck` / `emitDeck`) and
the editable surface are the stable surface, and the heuristic lane's inferred
model is expected to move.

## [Unreleased]

Nothing yet.

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

[unreleased]: https://github.com/shbernal/pptx-html/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/shbernal/pptx-html/releases/tag/v0.1.0
