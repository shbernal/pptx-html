# AGENTS.md

Guidance for coding agents working in the `dom2pptx/` project.

## Repository Expectations

- This repo builds `dom2pptx`: it turns HTML from the DOM into editable PPTX by
  driving `@shbernal/ts-pptx`. Its role is the **HTML → ts-pptx** link only.
- This is a **standalone top-level project** (`~/dev/dom2pptx`) with its own git.
  It is not part of a workspace; run its commands from this directory.
- Use `pnpm`. Node `>=24`. Keep source in `src/`, tests in `test/`. Treat `dist/`
  as generated build output.
- Preserve unrelated dirty state. Do not revert user changes.

## Purpose

Three goals, in priority order:

1. **Web previews.** HTML is the preview medium for a deck; this library is what
   makes that preview redeemable as a real `.pptx`.
2. **AI-authored decks.** Agents emit HTML well and OOXML badly. This is the
   adapter that makes agent-emitted HTML land as editable slides.
3. **(Deferred) Round-tripping.** PPTX in, HTML/DOM preview out. Not
   implemented; do not build toward it without being asked.

## Fidelity Is Heuristic

HTML/CSS and PPTX are different formats and neither is a superset of the other.
The mapping is best-effort by design:

- Aim for the closest **editable** PowerPoint construct, not a pixel match.
- When a faithful mapping is impossible, degrade and record a `Warning`. Never
  silently drop content.
- `convertDeckRaster` is the escape hatch for callers who want pixel fidelity
  instead of editability. Keep it working; do not let it rot.
- Do not add deterministic-looking guarantees to the docs. Fidelity claims must
  stay honest.

## Architecture & Boundaries

One-way dependency: **consumer app → dom2pptx → ts-pptx**. No cycles.

- `dom2pptx` does **not** know about AI or UI. It owns `ts-pptx` and
  `html2canvas` as dependencies; consumers never import `ts-pptx` directly.
- The IR (`src/ir/model.ts`) is the boundary between the two internal layers:
  - `src/extract/` (DOM → IR) is **browser-only** (iframe, `getComputedStyle`,
    `getBoundingClientRect`, canvas, fonts).
  - `src/emit/` (IR → ts-pptx) is **pure and isomorphic** — unit-testable
    without a browser. This is where custGeom correctness lives.
- Keep the IR as the single source of truth for the slide model. Changing the IR
  shape touches both layers; do it deliberately.

## Scope

- **In scope:** a documented subset of HTML/CSS aimed at slide layouts (sectioned
  slides, utility-class styling, iconify, gradients, tables, lists). Expand it by
  adding fixtures, not by chasing arbitrary CSS.
- **Out of scope:** rendering arbitrary web pages. This is not a browser.

## Fix Upstream When Possible

- Prefer fixing generic OOXML / emitter problems upstream in `@shbernal/ts-pptx`
  (it helps every consumer) over patching them here. When a fix belongs upstream
  but is not yet released, keep any stopgap here thin and clearly marked, and
  drop it once a release carries the fix.
- Anything discovered while building the custGeom/SVG-path vectorizer (a missing
  custGeom case, a measure gap) goes upstream, not patched locally.

## Testing

- **Unit (`test/unit/`, no DOM, fast):** feed hand-written `SlideModel` fixtures
  into `emit/*`, write to base64, and parse back with ts-pptx's `read`/`inspect`
  to assert structure. Prefer structural assertions over binary goldens.
- **Browser/e2e (`test/browser/`, Playwright + headless Chromium):** load HTML
  slide fixtures and run `convertSlide`/`convertDeck` with `output:'base64'` and a
  static injected `resolveIcon`. jsdom/happy-dom are insufficient.
- Keep most assertions in the headless unit layer; minimize the e2e surface.

## Verification

- For source changes, run `pnpm run build` and `pnpm run typecheck`.
- For behavior changes, run `pnpm run test:unit`.
- Lint and format with `pnpm run lint` / `pnpm run format:check`.

## Reference Order

Most questions here are answered by ts-pptx's own API surface, not by the OOXML
spec — this package drives a writer, it does not emit XML. Check in this order:

1. **ts-pptx's shipped type declarations.** The package publishes its API as
   `.d.ts` files next to `dist/`. Locate it with
   `node -e "console.log(require.resolve('@shbernal/ts-pptx/package.json'))"`
   (pnpm hides the real directory under `node_modules/.pnpm/`), then read the
   `exports` map. The subpaths that matter here:
   - `@shbernal/ts-pptx` — the writer: `addShape` / `addText` / `addTable` /
     `addImage` option shapes, `ShapeType.custGeom`, and the freeform point DSL
     that `src/emit/custgeom.ts` passes through unchanged.
   - `@shbernal/ts-pptx/read` — `Presentation.load`, the round-trip oracle the
     unit tests already assert against.
   - `@shbernal/ts-pptx/inspect` — per-element view (box, fill, text runs,
     paragraph boundaries, `a:bodyPr` autofit mode). Use it to confirm what an
     emit path or a repair actually produced.
   - `@shbernal/ts-pptx/measure` — font metrics and measured text fit, for any
     mapping decision that depends on whether text will fit its frame.
   - `@shbernal/ts-pptx/zip` — the fflate ZIP toolkit `src/repair/` uses. Do not
     add a separate ZIP dependency.
2. **The `ooxml` MCP** (ECMA-376 schema/spec) — for raw XML only, see below.
3. **Web search** — last resort.

Do not vendor large spec text into the repo.

## Raw OOXML Work

`src/repair/repair.ts` is the only code here that touches OOXML directly,
rewriting `ppt/slides/slideN.xml`, `presentation.xml`, `[Content_Types].xml`,
and the `.rels` parts after the writer has run. Everything else goes through
ts-pptx's DSL and should stay that way.

- Consult the `ooxml` MCP before changing any of those transforms. Children of
  `a:bodyPr` are a schema-ordered sequence, attributes like `p:sldSz@type` are
  enumerations, and dropping relationship entries can leave a package
  PowerPoint refuses to open.
- Verify, do not eyeball: write a deck, run the repair, read it back with
  `read` / `inspect` in a unit test. The file is `@ts-nocheck` and the
  transforms are regex over XML, so nothing else will catch a mistake.
- `repairPptxBase64` catches every failure and returns the *unrepaired* deck, so
  a broken repair degrades silently. That is why the round-trip assertion is not
  optional.
- These repairs are stopgaps for ts-pptx bugs (autofit defaults, duplicate
  `cNvPr` ids, generated notes masters, missing slide-size type). Per **Fix
  Upstream When Possible**, prefer moving one upstream to extending it here, and
  delete the local repair once a release carries the fix.
