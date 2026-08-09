---
layout: home
editLink: false
hero:
  name: 'pptx-html'
  text: 'Edit a PowerPoint deck as a web page'
  tagline: 'and get the deck back — not an approximation of it.'
  actions:
    - theme: brand
      text: Playground
      link: /playground
    - theme: alt
      text: Read the design record
      link: /docs/
features:
  - title: Modeled
    details: The intermediate representation represents it, and it survives the loop exactly.
  - title: Carried
    details: The IR does not model it, so its XML moves across untouched. No approximation, and no loss.
  - title: Warned
    details: It can be neither modeled nor carried, and the conversion says so. A visible failure, never a silent one.
---

## The loop

`pptx-html` reads a `.pptx` into a slide model, renders that model as HTML, reads
the edited HTML back, and writes a `.pptx` out again. Four legs, one loop:

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
canonicalized before diffing.

## The fourth state, ruled out

Modeled, carried and warned are the three states above. What they exist to rule
out is the fourth — *approximated*: content that comes out looking about right
but has no way back. It is why the raster fallback was removed rather than kept
as an escape hatch. A slide flattened into a picture is the one output that can
never re-enter the loop, so producing a file that way is a failure wearing a
success's clothes.

That is also why this site shows no pictures of slides. A screenshot of a deck
would prove exactly the thing the project refuses to do. Anything this site shows
of a deck has to be produced by running the library.

## What the guarantee covers

A generated corpus of 17 decks runs the full loop on every CI build, and the
per-construct fidelity ledger is snapshotted so it cannot move silently. What
that gate currently covers:

- **Input domain** — decks written by
  [`@shbernal/ts-pptx`](https://www.npmjs.com/package/@shbernal/ts-pptx), which is
  what the generated corpus is made of. Decks authored in PowerPoint are a
  deliberate second tier and are **not yet gated**.
- **Environment** — the loop is host-agnostic and runs in Node and the browser
  alike. The best-effort heuristic lane, for HTML this library did not render, is
  browser-only: it needs a real DOM.
- **Distribution** — published to npm as `pptx-html`. ESM only, Node `>=24`.
  Pre-1.0: the loop's four legs are stable, the heuristic lane's model is not.

Claims here are held to what the oracle actually gates. Where something is not
covered, it says so.

## Start here

- [Invariant R and the round-trip oracle](/docs/round-trip) — the property, and
  the harness that turns it from a claim into a gate.
- [Architecture](/docs/architecture) — the two lanes, the two models, and the
  rules each part of the loop is built on.
- [Decisions that must not be undone](/docs/decisions) — two capabilities that
  were deliberately removed, and why adding them back would cost more than it
  looks.
- [README](https://github.com/shbernal/pptx-html#readme) — the introduction, with
  install and the four-leg example.
