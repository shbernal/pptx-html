# Changelog

Notable changes per release. Format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/); versions follow
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

Pre-1.0, the two lanes carry different promises and the version reflects that:
the loop's four legs (`importDeck` / `renderDeck` / `parseDeck` / `emitDeck`) and
the editable surface are the stable surface, and the heuristic lane's inferred
model is expected to move.

## [Unreleased]

## [0.1.0] — 2026-08-09

First public release.

### Added

- **The loop.** `importDeck`, `renderDeck`, `parseDeck` and `emitDeck` — `.pptx`
  to a paint model, to HTML, back to the model, back to `.pptx`. Anything the
  model does not express is carried across as its original XML rather than
  approximated, and anything that can be neither modeled nor carried is warned
  about.
- **Invariant R and the oracle that gates it.** A generated corpus of 17 decks
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
- **Lane reporting.** `parseDeck` states per slide how it read the document
  (`exact`, `reconciled`, `drifted`, `heuristic`) instead of leaving a caller to
  infer success from the fact that a file appeared.
- **The heuristic lane.** `convertDeck` / `convertSlide` infer a model from a
  rendered DOM, for HTML this library did not produce. Browser-only,
  explicitly best-effort, and sharing no code path with the loop.
- **A runnable example.** `examples/round-trip.mjs` runs all four legs, edits a
  run, and re-imports the emitted deck to show the edit arrived.

### Notes on scope

- The guarantee covers decks written by `@shbernal/ts-pptx`, which is what the
  corpus is made of. Decks authored in PowerPoint are a deliberate second tier
  and are not yet gated.
- There is no rasterizer, on purpose. A flattened slide is the one output that
  can never re-enter the loop — see
  [docs/decisions.md](./docs/decisions.md).

[unreleased]: https://github.com/shbernal/pptx-html/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/shbernal/pptx-html/releases/tag/v0.1.0
