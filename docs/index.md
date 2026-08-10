---
doc-schema-version: 1
title: 'pptx-html docs'
summary: 'Index of the design record: what the project guarantees, how it is built, and what was deliberately left out.'
read_when:
  - Starting work on this repo
  - Looking for the rationale behind a design constraint
doc_type: 'overview'
---

# pptx-html docs

`pptx-html` moves slides between HTML and PPTX by driving
[`@shbernal/ts-pptx`](https://www.npmjs.com/package/@shbernal/ts-pptx). It owns
the **HTML ⇄ ts-pptx** link, in both directions. It does not emit OOXML itself,
and it does not generate the HTML.

These pages are the design record: the *why*. The [README](../README.md) is the
introduction, and [CONTRIBUTING](../CONTRIBUTING.md) is how to build, test and
verify a change.

| Page | What it answers |
| --- | --- |
| [Invariant R and the round-trip oracle](./round-trip.md) | What the project guarantees, and the harness that gates it |
| [Architecture](./architecture.md) | The two lanes, the two models, and the rules each part is built on |
| [Decisions that must not be undone](./decisions.md) | Two capabilities that were removed on purpose |

## The shape of it in one screen

```text
.pptx  ──import──►  IR  ──render──►  HTML   (what a human sees / edits)
  ▲                  ▲                 │
  └────emit──────────┴─────parse───────┘   (what a machine reads back)
```

Four legs, one loop, over a paint model that is embedded in the rendered document
as a JSON island. The return path parses that island, never `getComputedStyle`,
which is what lets the visible SVG approximate freely without any of that
approximation reaching the emitted deck.

A second, secondary lane (`convertDeck` / `convertSlide`) infers a model from a
rendered DOM, for HTML this library did not produce. It is browser-only,
best-effort, and shares no code path with the loop.
