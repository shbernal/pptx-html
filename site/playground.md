---
title: 'Playground'
description: 'Run the round-trip loop in your own browser: import, render, edit, emit, download.'
editLink: false
aside: false
pageClass: pxh-wide
---

# Playground

`importDeck → renderDeck → parseDeck → emitDeck`, running here, on this page, in
your browser. Not a recording of a conversion and not a picture of a slide: the
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

The panel beside the preview is the **editable surface**: run text, `bold`,
`italic`, `underline`, `strike`, `sizePt`, `color`, a paragraph's `align`,
`bullet`, `marginLeftPt` and `indentPt`, and deleting a node. That list is not a
subset chosen for the demo: it is `EDITABLE_SURFACE`, exported from the package,
and there is deliberately no control for anything outside it. An input whose
value was silently dropped on the way back would be the exact failure this
project is built to refuse.

`underline`, `strike`, `align` and `bullet` are drop-downs rather than checkboxes,
and the extra option is the point: **inherited** clears the property, while
**none** states outright that the run is not underlined. A run that would
otherwise take an underline from its placeholder needs the second, and a checkbox
has no way to say it.

`bullet` is the drop-down that took an upstream change to exist. Until
[ts-pptx#15](https://github.com/shbernal/ts-pptx/issues/15), the write API could
say *this paragraph has a bullet* and *this paragraph has none* but not *this
paragraph says nothing about its bullet* (the omitted option wrote the explicit
"none"), so an **inherited** position would have produced the wrong one of the
three, and the two look identical on screen. Rather than ship a control that lies
in a way nothing on the page could show, there was no control at all until the
option gained an `inherit` spelling.

The two margin fields (`marginLeftPt` is where the body text starts, `indentPt`
how far the first line sits from it) came with the same upstream change, and an
empty field is the drop-downs' **inherited** in another shape: a
paragraph whose margin is cleared follows its list style again, which is not the
same paragraph as one stating `0`. Until `bullet` gained its third state the
bullet decided both attributes, so *suppress this bullet but keep the margin it
inherits* was not a deck the write API could produce.

Its fourth position, *as the deck states it*, is disabled and is the same rule
still doing its job. A paragraph can hold a numbered bullet, a picture bullet or a
glyph with its own colour, and the write API cannot author every one of those
back, so the panel says the deck states one and declines to replace it, instead
of showing "none" beside a visible bullet or rounding a numbering scheme to a dot.
Nothing is lost by leaving it alone: an untouched bullet is never rewritten.

## What it does not prove

The round-trip oracle gates decks written by
[`@shbernal/ts-pptx`](https://www.npmjs.com/package/@shbernal/ts-pptx); the
samples above are that corpus. A deck you author in PowerPoint and drop here runs
the same code, but it is the project's **second tier** and is not gated by CI. If
one comes back with warnings or on a weaker lane, that is the documented state of
the work, and showing it is the point.

[Invariant R and the round-trip oracle](/docs/round-trip) is the full statement of
what is and is not guaranteed.
