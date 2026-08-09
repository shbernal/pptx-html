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

pnpm run example      # the loop end to end; needs `build` first

pnpm run site:dev     # the public site, locally
pnpm run site:build
pnpm run site:preview # serve the built site at its real base path
```

## The site

[`site/`](./site) builds <https://shbernal.github.io/pptx-html/>. It owns the
build; `docs/` stays the tracked design record and is read as input.

**`site/docs/` is generated — never edit it.** `site/scripts/sync-docs.ts` mirrors
`docs/*.md` into it, translating the frontmatter and rewriting the links that
point at files the site does not publish, and
`site/scripts/build-ledger.ts` adds the fidelity ledger by running the oracle's
coverage reporter over the corpus. Both run from `site:prepare`, which `site:dev`
and `site:build` call first. Edit `docs/`.

Those scripts fail rather than warn, on purpose: a page missing from
`docs/docs.json`, a link leaving `docs/` with no mapping, an unrecognised
frontmatter key, or a playground sample naming a corpus deck that no longer
exists all stop the build. Drift between the record and the site is the failure
they exist to make impossible, and they can only deliver that if it is an error.

`site:dev` and `site:build` run `pnpm run build` first. The site imports
`pptx-html` **by name**, through the `exports` map, exactly as a consumer would —
the same reason `examples/round-trip.mjs` does — so `dist/` has to exist.

The home page (`site/index.md`) is a landing page rather than a doc page, and the
two rows of slides drifting across it are **rendered on load, in the visitor's
browser**, by `site/.vitepress/theme/home/`. `decks.ts` writes two eight-slide
consulting decks with `@shbernal/ts-pptx`, `showcase.ts` runs
`importDeck → renderDeck` and cuts the document into its `section.pxh-slide`
elements verbatim, and each one is installed in a shadow root so the site's own
stylesheet cannot restyle it. They are **not** corpus decks and must not become
them: the corpus is the oracle's input domain, and widening it to decorate a page
would change what the gate means. They are decks the writer can write, which is
the domain the guarantee is scoped to, and the page says so under the rows.

Two rules for anything under `site/**`:

- **No claim the oracle does not gate.** The site is the project's marketing
  surface and therefore the likeliest place to overstate. Every page that touches
  fidelity states the input domain — decks written by `@shbernal/ts-pptx` — and
  says that PowerPoint-authored decks are a deliberate second tier.
- **Nothing is shown that was not produced by running the library.** There are no
  slide images anywhere on the site. A screenshot of a converted deck is exactly
  the impression [the raster decision](./docs/decisions.md) exists to refuse, so
  previews are rendered by `renderDeck` in the visitor's browser or they are not
  shown at all.

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
| `docs/`, or anything under `site/` | `pnpm run site:build` — dead-link checking is on, and it is the acceptance test for the docs mirror |
| Anything at all | `pnpm run check` |

**After changing what `src/index.ts` exports, read `dist/index.d.ts`.** The public
surface has a failure mode neither `tsc` nor the tests can see: nothing in-repo
imports through the entry point, and `export type *` from a module that also
exports values emits those values as *types named after functions* — declared,
uncallable, and discovered only by a consumer. `examples/round-trip.mjs` is the
one thing that imports the package *by name*, so it catches the coarse version of
this — a broken `exports` map — but not the shape of a type. CI runs it; the
reading is still on you.

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
is deleted. There is no live one to copy right now — 3.1.0 retired the last of
them — so `test/oracle/script-lane.ts` and `site/.vitepress/theme/home/decks.ts`
show the shape at its other end: what the comment turns into once the release
lands, which is a sentence about why the deleted stopgap existed rather than a
deletion nobody can date.

`ts-pptx-upstream`, the skill the package ships, is the normative reference for
writing the report itself. See [AGENTS.md](./AGENTS.md#upstream) for what belongs
upstream and what stays here.

## Commits

Conventional-commit subjects (`feat:`, `fix:`, `docs:`, `refactor:`, `chore:`),
scoped where it helps (`feat(parse):`). Write the subject as what the change does
for a reader of the repo, not as a plan step.

`pnpm run check` must be clean; a `lefthook` pre-commit job runs
`biome check --write` on staged files if hooks are wired in your environment.
