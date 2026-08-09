# Contributing

Read [docs/](./docs/index.md) before changing anything structural — most of the
constraints in this codebase are load-bearing and the reasoning is written down.

## Setup

Requires Node `>=24` and pnpm.

```bash
pnpm install
```

The writer dependency, `@shbernal/ts-pptx`, is consumed from public npm — no local
link or sibling checkout is required.

`lefthook`'s postinstall is denied on purpose in `pnpm-workspace.yaml`, and there
is deliberately no `prepare: lefthook install` script. That postinstall syncs git
hooks, and when `core.hooksPath` is set globally it writes *into that global
directory*, replacing the existing `pre-commit` with a wrapper hardcoded to this
repo's `node_modules` — which would then fire for every repo on the machine.
`lefthook run` needs no installed hooks, so hook wiring is left to whatever the
environment already provides.

## Commands

```bash
pnpm run build        # tsdown → ESM dist/
pnpm run typecheck    # tsc --noEmit
pnpm run check        # biome: lint + format + import sorting
pnpm run check:fix

pnpm run test:unit    # node, no DOM, fast
pnpm run test:oracle  # the round-trip gate, no DOM
pnpm run test:browser # Playwright + headless Chromium
pnpm run test         # build, then all three
```

## The three test layers

Each runs in the environment it needs; the split is enforced by vitest projects.

**Unit (`test/unit/`, no DOM, fast).** The pure pieces of both lanes — the paint
model and its surface, the island, the reconcile fold and the extractor boundary;
plus hand-written `heuristic/` fixtures written to base64 and parsed back with
ts-pptx's `read`/`inspect`. Prefer structural assertions over binary goldens.

**Oracle (`test/oracle/`, no DOM).** The round-trip gate. See
[docs/round-trip.md](./docs/round-trip.md) — its rules are normative, not
stylistic.

**Browser (`test/browser/`, Playwright + headless Chromium).** The two things that
cannot be faked: reading the editable surface back out of a rendered document, and
the heuristic lane's extract path (`convertSlide`/`convertDeck` with
`output:'base64'` and a static injected `resolveIcon`). jsdom and happy-dom are
insufficient — they have no real layout or canvas.

Keep most assertions in the headless layers and minimize the e2e surface.

**Assert on the lane a slide took, not just on the file that came out.** A lane
that degrades gracefully hides its own bugs; three past return-path bugs each
produced a plausible, working document while quietly doing nothing. See the end of
[docs/round-trip.md](./docs/round-trip.md).

## What to run for which change

| Change | Run |
| --- | --- |
| Any source change | `pnpm run build`, `pnpm run typecheck` |
| Behaviour | `pnpm run test:unit` |
| Anything touching fidelity | `pnpm run test:oracle`, and update the coverage snapshot deliberately — never a blind `-u` |
| Renderer, surface reading, or the heuristic lane | `pnpm run test:browser` |
| Anything at all | `pnpm run check` |

**After changing what `src/index.ts` exports, read `dist/index.d.ts`.** The public
surface has a failure mode neither `tsc` nor the tests can see: nothing in-repo
imports through the entry point, and `export type *` from a module that also
exports values emits those values as *types named after functions* — declared,
uncallable, and discovered only by a consumer.

## Where to look things up

Most questions here are answered by ts-pptx's own API surface, not by the OOXML
spec — this package drives a writer, it does not emit XML. Check in this order:

1. **ts-pptx's shipped type declarations.** The package publishes its API as
   `.d.ts` files next to `dist/`. Locate it with
   `node -e "console.log(require.resolve('@shbernal/ts-pptx/package.json'))"`
   (pnpm hides the real directory under `node_modules/.pnpm/`), then read the
   `exports` map. The subpaths that matter here:
   - `@shbernal/ts-pptx` — the writer: `addShape` / `addText` / `addTable` /
     `addImage` option shapes, `ShapeType.custGeom`, and the freeform point DSL
     that `src/heuristic/custgeom.ts` passes through unchanged.
   - `@shbernal/ts-pptx/read` — `Presentation.load`, the round-trip oracle the
     unit tests already assert against.
   - `@shbernal/ts-pptx/inspect` — per-element view (box, fill, text runs,
     paragraph boundaries, `a:bodyPr` autofit mode). Use it to confirm what an
     emit path actually produced.
   - `@shbernal/ts-pptx/measure` — font metrics and measured text fit. Read-only,
     and **not from the emit path**: a measurement taken at emit time makes the
     same IR produce different decks on different machines. Measuring belongs in
     `src/heuristic/`, resolved into that lane's model where it is taken.
   - `@shbernal/ts-pptx/zip` — the fflate ZIP toolkit, if a test ever needs to look
     inside a package. Do not add a separate ZIP dependency.
2. **The `ooxml` MCP** (ECMA-376 schema/spec) — for raw XML questions only. Note
   that [there is no raw OOXML work in `src/`](./docs/decisions.md), so this is
   almost always for understanding, not for writing.
3. **Web search** — last resort.

Do not vendor large spec text into the repo.

## Fix upstream when possible

Prefer fixing generic OOXML / emitter problems upstream in `@shbernal/ts-pptx` —
it helps every consumer — over patching them here. When a fix belongs upstream but
is not yet released, keep any stopgap here thin and clearly marked, and drop it
once a release carries the fix.

**File the issue in the same unit of work that found the gap**, on
`shbernal/ts-pptx`, before the commit — not batched into a later sweep. A gap that
lives only in a chat session or a scratch doc does not exist: sessions end and
scratch plans are deleted by design, and the tracker is the one queue that
survives both. An ask that turns out to be unclear gets closed, which is cheap; a
gap never filed costs a rediscovery, usually by whoever re-invents the same
workaround.

Every stopgap therefore carries its **issue URL** and the condition under which it
is deleted — see `src/import/paint.ts` and `test/oracle/script-lane.ts` for the
shape of that comment.

## Commits

Conventional-commit subjects (`feat:`, `fix:`, `docs:`, `refactor:`, `chore:`),
scoped where it helps (`feat(parse):`). Write the subject as what the change does
for a reader of the repo, not as a plan step.

`pnpm run check` must be clean; a `lefthook` pre-commit job runs
`biome check --write` on staged files if hooks are wired in your environment.
