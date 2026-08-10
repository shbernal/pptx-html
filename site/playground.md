---
title: 'Playground'
description: 'Run the round-trip loop in your own browser: import, render, edit, emit, download.'
editLink: false
aside: false
pageClass: pxh-wide
---

# Playground

`importDeck → renderDeck → parseDeck → emitDeck`, running here, on this page, in
your browser. Not a recording of a conversion and not a picture of a slide — the
same four functions a consumer calls, on bytes that never leave the tab.

<ClientOnly>
	<Playground />
</ClientOnly>

## What you are looking at

The four boxes are the loop's four legs, named for the functions themselves. Each
one reports what it produced and how long it took, and a leg that throws goes red
while the ones after it say they did not run. There is no state in that strip that
was not earned.

The **lane** under each slide is the return path's own account of how it read the
document. `exact` means the model came back as it was rendered; `reconciled` means
a sanctioned edit was folded in; `drifted` means something outside the editable
surface changed and the island's value stood instead. It is reported per slide,
because a deck where one slide was edited and eleven were not should not describe
all twelve the same way.

The panel beside the preview is the **editable surface** — run text, `bold`,
`italic`, `underline`, `strike`, `sizePt`, `color`, and deleting a node. That list
is not a subset chosen for the demo: it is `EDITABLE_SURFACE`, exported from the
package, and there is deliberately no control for anything outside it. An input
whose value was silently dropped on the way back would be the exact failure this
project is built to refuse.

`underline` and `strike` are drop-downs rather than checkboxes, and the extra
option is the point: **inherited** clears the property, while **none** states
outright that the run is not underlined. A run that would otherwise take an
underline from its placeholder needs the second, and a checkbox has no way to say
it.

## What it does not prove

The round-trip oracle gates decks written by
[`@shbernal/ts-pptx`](https://www.npmjs.com/package/@shbernal/ts-pptx) — the
samples above are that corpus. A deck you author in PowerPoint and drop here runs
the same code, but it is the project's **second tier** and is not gated by CI. If
one comes back with warnings or on a weaker lane, that is the documented state of
the work, and showing it is the point.

[Invariant R and the round-trip oracle](/docs/round-trip) is the full statement of
what is and is not guaranteed.
