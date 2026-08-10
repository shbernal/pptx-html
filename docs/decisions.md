---
doc-schema-version: 1
title: 'Decisions that must not be undone'
summary: 'Two capabilities were deliberately removed; both look like obvious things to add back.'
read_when:
  - Considering a rasterizer, a preview image or a slide-to-PNG path
  - Considering a post-write fixup that rewrites OOXML
  - Wondering why an obvious-looking feature is missing
doc_type: 'decision'
---

# Decisions that must not be undone

Both entries below describe code that existed, was measured, and was deleted.
Both are the kind of thing a newcomer proposes in their first week.

## The raster path was removed on purpose

`convertDeckRaster` / `rasterizeSlide` and the `html2canvas` dependency are gone,
and **nothing should reintroduce a rasterizer in `src/`**.

A slide flattened into a full-bleed picture has no model to re-import: it is the
one output that can never re-enter the loop, so it is a failure wearing a
success's clothes. The cost is real and was accepted: hostile HTML no longer
always yields *some* file. A warned failure beats a file that can never come back.

The obvious follow-up, "then render `.pptx` → PNG directly", was asked and
answered: it cannot run in a browser. It needs a PowerPoint-grade renderer, and
the slides that would most want a preview are the carried ones, which by
definition have no model to draw from. The real options (PowerPoint COM
`Slide.Export`, LibreOffice headless) are out-of-process and host-dependent, so a
preview generator is a **test utility, not a package feature**. One under `test/`
for visual review of the corpus is welcome; nothing goes in `src/`.

Canvas use that remains in `src/heuristic/` is *not* this: it rasterizes
individual CSS gradients and re-encodes images, and each of those is a modeled
element. It is not slide rasterization.

## There is no raw OOXML work, and there must not be another repair layer

Everything goes through ts-pptx's DSL.

`src/repair/repair.ts` used to rewrite `ppt/slides/slideN.xml`,
`presentation.xml`, `[Content_Types].xml` and the `.rels` parts after the writer
had run. It was measured against ts-pptx 3.0.0 across the whole corpus and
deleted. Four of its five rules never fired, and both that did were wrong: one
flattened an inherited autofit into an explicit one, the other **deleted every
speaker note in the deck**. It also swallowed exceptions and returned the
unrepaired file, so its output depended on whether something threw.

The lesson is worth more than the code was: **a layer that silently "fixes" things
hides both its own obsolescence and its own damage.** Every one of those defects
was invisible from outside: the deck opened, the text was there, and only the
notes were gone.

`test/oracle/writer-output.test.ts` now asserts the conditions the module existed
for do not occur. If the writer regresses, the answer is an upstream issue and a
failing test, **not a second repair pass**. If you find yourself reaching for a
post-write rewrite, read that file first: it names each rule and why it went.
