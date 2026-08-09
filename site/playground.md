---
title: 'Playground'
description: 'Run the round-trip loop in your own browser.'
editLink: false
---

# Playground

**Not built yet.** This page is a placeholder, and it says so rather than showing
a mock-up of what it will be — the same rule the library applies to a slide it
cannot convert.

What will be here: `importDeck → renderDeck → parseDeck → emitDeck` running in
your browser, on a sample deck, with the rendered HTML editable through the
surface the library declares and the resulting `.pptx` downloadable. Not a
picture of a deck, and not a server round-trip — the actual loop, in the page.

Until then, the loop is runnable in a checkout:

```bash
pnpm install && pnpm run build && pnpm run example
```

[`examples/round-trip.mjs`](https://github.com/shbernal/pptx-html/blob/main/examples/round-trip.mjs)
imports a deck, edits a run, writes the deck back out, and re-imports it to show
the edit arrived.
