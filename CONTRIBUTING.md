# Contributing

Read [docs/](./docs/index.md) before changing anything structural. Most of the
constraints in this codebase are load-bearing and the reasoning is written down.

## Setup

Requires Node `>=24` and pnpm.

```bash
pnpm install
```

The writer dependency, `@shbernal/ts-pptx`, is a **released version from npm**.
The range lives in `package.json` and is not repeated here; it moves often enough
that a second copy in prose is a copy that goes stale, so read it from the
manifest. It was pinned to a git sha for as long as the fixes this repo depends on
were unreleased. That is over, and a range is what a published `pptx-html` has to
carry anyway: a git URL in `dependencies` travels to consumers and would make
every installer build the writer from source, which `prepublishOnly` does not
check for.

To move to a newer release:

```bash
pnpm add "@shbernal/ts-pptx@^<version>"
npx skills add ./node_modules/@shbernal/ts-pptx -s '*' -a claude-code -a codex -a universal -y
pnpm run test:oracle
```

The skill reinstall is part of the bump, not a separate chore: the skill ships
inside the package, so it moves with the version ([AGENTS.md](./AGENTS.md)).

Two things the sha era left behind in `pnpm-workspace.yaml`, both still live:

- `'@shbernal/ts-pptx': false` in `allowBuilds`. The entry is no longer about
  building the writer, since the published tarball ships `dist/`, but the manifest
  still declares `prepare`, which pnpm counts as a build script and asks about for
  a registry dependency too. Deleting the line makes pnpm write `set this to true
  or false` back into the file; `false` is correct because that `prepare` is
  upstream's own dev wiring and does nothing for a consumer.
- `minimumReleaseAgeExclude`. Inert while the pin was a sha, because the gate reads
  a registry publish date and a git tarball has none. It is now the thing that
  keeps a same-day bump installable.

If a fix is needed before it is released, `npm i github:shbernal/ts-pptx#<sha>`
works. That install builds from source and pulls upstream's `devDependencies`, so
it is for trying an unreleased fix, not for staying on. Going back to a sha means
restoring `allowBuilds` to `true`.

`lefthook`'s postinstall is denied on purpose in `pnpm-workspace.yaml`, and there
is deliberately no `prepare: lefthook install` script. That postinstall syncs git
hooks, and when `core.hooksPath` is set globally it writes *into that global
directory*, replacing the existing `pre-commit` with a wrapper hardcoded to this
repo's `node_modules`, which would then fire for every repo on the machine.
`lefthook run` needs no installed hooks, so hook wiring is left to whatever the
environment already provides.

## Commands

```bash
pnpm run build        # tsdown → ESM dist/
pnpm run typecheck    # tsc --noEmit, over src/, test/ and site/; needs `build` first
pnpm run check        # oxlint + oxfmt: lint, format and import sorting
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

**`site/docs/` is generated. Never edit it.** `site/scripts/sync-docs.ts` mirrors
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
`pptx-html` **by name**, through the `exports` map, exactly as a consumer would,
which is the same reason `examples/round-trip.mjs` does, so `dist/` has to exist.

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
  fidelity states the input domain, decks written by `@shbernal/ts-pptx`, and
  says that PowerPoint-authored decks are a deliberate second tier.
- **Nothing is shown that was not produced by running the library.** There are no
  slide images anywhere on the site. A screenshot of a converted deck is exactly
  the impression [the raster decision](./docs/decisions.md) exists to refuse, so
  previews are rendered by `renderDeck` in the visitor's browser or they are not
  shown at all.

## The three test layers

Each runs in the environment it needs; the split is enforced by vitest projects.

**Unit (`test/unit/`, no DOM, fast).** The pure pieces of both lanes: the paint
model and its surface, the island, the reconcile fold and the extractor boundary;
plus hand-written `heuristic/` fixtures written to base64 and parsed back with
ts-pptx's `read`/`inspect`. Prefer structural assertions over binary goldens.

**Oracle (`test/oracle/`, no DOM).** The round-trip gate. See
[docs/round-trip.md](./docs/round-trip.md); its rules are normative, not
stylistic.

**Browser (`test/browser/`, Playwright + headless Chromium).** The two things that
cannot be faked: reading the editable surface back out of a rendered document, and
the heuristic lane's extract path (`convertSlide`/`convertDeck` with
`output:'base64'` and a static injected `resolveIcon`). jsdom and happy-dom are
insufficient: they have no real layout or canvas.

Keep most assertions in the headless layers and minimize the e2e surface.

**Assert on the lane a slide took, not just on the file that came out.** A lane
that degrades gracefully hides its own bugs; three past return-path bugs each
produced a plausible, working document while quietly doing nothing. See the end of
[docs/round-trip.md](./docs/round-trip.md).

## What to run for which change

| Change | Run |
| --- | --- |
| Any source or test change | `pnpm run build`, `pnpm run typecheck` |
| Behaviour | `pnpm run test:unit` |
| Anything touching fidelity | `pnpm run test:oracle`, and update the coverage snapshot deliberately, never a blind `-u` |
| Renderer, surface reading, or the heuristic lane | `pnpm run test:browser` |
| `docs/`, or anything under `site/` | `pnpm run site:build`; dead-link checking is on, and it is the acceptance test for the docs mirror |
| Anything at all | `pnpm run check` |

**After changing what `src/index.ts` exports, read `dist/index.d.ts`.** The public
surface has a failure mode neither `tsc` nor the tests can see: nothing in-repo
imports through the entry point, and `export type *` from a module that also
exports values emits those values as *types named after functions*: declared,
uncallable, and discovered only by a consumer. `examples/round-trip.mjs` is the
one thing that imports the package *by name*, so it catches the coarse version of
this, a broken `exports` map, but not the shape of a type. CI runs it; the
reading is still on you.

## Where to look things up

Most questions here are answered by ts-pptx's own types, not by the OOXML spec.
This package drives a writer; it does not emit XML. Check in this order:

1. **ts-pptx's shipped type declarations.** The package publishes its API as
   `.d.ts` files next to `dist/`. Locate it with
   `node -e "console.log(require.resolve('@shbernal/ts-pptx/package.json'))"`
   (pnpm hides the real directory under `node_modules/.pnpm/`), then read the
   `exports` map. The subpaths that matter here:
   - `@shbernal/ts-pptx`, the writer: `addShape` / `addText` / `addTable` /
     `addImage` option shapes, `ShapeType.custGeom`, and the freeform point DSL
     that `src/heuristic/custgeom.ts` passes through unchanged.
   - `@shbernal/ts-pptx/read`: `Presentation.load`, the round-trip oracle the
     unit tests already assert against.
   - `@shbernal/ts-pptx/inspect`: per-element view (box, fill, text runs,
     paragraph boundaries, `a:bodyPr` autofit mode). Use it to confirm what an
     emit path actually produced.
   - `@shbernal/ts-pptx/measure`: font metrics and measured text fit. Read-only,
     and **not from the emit path**: a measurement taken at emit time makes the
     same IR produce different decks on different machines. Measuring belongs in
     `src/heuristic/`, resolved into that lane's model where it is taken.
   - `@shbernal/ts-pptx/zip`: the fflate ZIP toolkit, if a test ever needs to look
     inside a package. Do not add a separate ZIP dependency.
2. **The `ooxml` MCP** (ECMA-376 schema/spec), for raw XML questions only. Note
   that [there is no raw OOXML work in `src/`](./docs/decisions.md), so this is
   almost always for understanding, not for writing. It is `mcp-server-ooxml`
   run locally over stdio, as in `ts-pptx`: the schema graph ships inside the
   package, so it needs no account and no network and answers deterministically.
   It carries the schema and nothing else: no spec prose, no OPC part or
   content-type catalogue, no Microsoft-proprietary behaviour. Those go to
   step 3.
3. **Web search**, last resort.

Do not vendor large spec text into the repo.

## Fix upstream when possible

Prefer fixing generic OOXML / emitter problems upstream in `@shbernal/ts-pptx`,
where it helps every consumer, over patching them here. The dependency is a
published range now rather than a sha, so a fix arrives on the release that
carries it, and a stopgap covers the window between the two. That window is
longer than it was during the sha era, which is an argument for keeping the
stopgap thin and clearly marked, not for skipping the upstream fix: a `github:`
install is available for *testing* one before it ships (see [Setup](#setup)), and
is not a state to stay in now that `pptx-html` is published.

**File the issue in the same unit of work that found the gap**, on
`shbernal/ts-pptx`, before the commit, not batched into a later sweep. A gap that
lives only in a chat session or a scratch doc does not exist: sessions end and
scratch plans are deleted by design, and the tracker is the one queue that
survives both. An ask that turns out to be unclear gets closed, which is cheap; a
gap never filed costs a rediscovery, usually by whoever re-invents the same
workaround.

Every stopgap therefore carries its **issue URL** and the condition under which it
is deleted. None is live at the moment; the last of them went out with the release
that fixed it. So `test/oracle/script-lane.ts` and
`site/.vitepress/theme/home/decks.ts` show the shape at its other end: what the
comment turns into once the fix lands, which is a sentence about why the deleted
stopgap existed rather than a deletion nobody can date.

`ts-pptx-upstream`, the skill the package ships, is the normative reference for
writing the report itself. See [AGENTS.md](./AGENTS.md#upstream) for what belongs
upstream and what stays here.

## Releasing

Publishing runs from CI, through npm's **trusted publishing**. The job in
[`.github/workflows/publish.yml`](./.github/workflows/publish.yml) mints a
short-lived OIDC token, npm exchanges it for a credential scoped to this package
and to that one workflow file, and the package comes out carrying provenance
that points back at the run. There is no `NPM_TOKEN` in the repo secrets, and
there should not be one: a stored credential outlives every job that uses it and
is worth stealing, which a token that expires in minutes and cannot publish
anything else is not.

Two things are load-bearing and easy to break while tidying:

- **The filename.** npm's trusted-publisher configuration names `publish.yml`.
  Renaming the file revokes the package's ability to publish until the
  configuration on npmjs.com is changed to match, and the failure arrives at
  publish time, not at lint time.
- **The trigger.** Publishing the GitHub *release* is what starts it, not pushing
  the tag. A tag travels on a `git push --follow-tags` and can be cut by a script
  or by accident; publishing a release is a decision, and it is also the act that
  produces the notes a consumer reads. So the deliberate half stays human and the
  irreversible half follows it.

The workflow re-checks that the tag matches `package.json` and that the changelog
has a section for the version, then runs `pnpm run test`, the browser layer
included, and the example before publishing. `prepack` builds and
`prepublishOnly` runs typecheck, the lint/format check and the oracle at the
moment of publishing, exactly as they would from a laptop, so the two paths
cannot drift apart.

### Before you tag

What no check can have an opinion about:

1. **Bump `version` in `package.json`.** Pre-1.0 the loop's four legs and the
   editable surface are the stable surface and the heuristic lane's model is not,
   so a break in the latter is a minor rather than the major it will be after 1.0.
2. **Move `[Unreleased]` into a dated section in `CHANGELOG.md`** and update both
   link references at the bottom of the file. Write it for someone deciding
   whether to upgrade, not as a commit log.
3. **Read `dist/index.d.ts`.** The one check nothing automates; see the note
   above about `export type *`.
4. **Inspect the tarball**: `npm pack --dry-run`. `files` is `dist` plus the
   changelog, and npm adds `README.md`, `LICENSE` and `package.json` on its own;
   anything else appearing is a bug in `files`.

Commit that and let CI go green on `main`. The site is deployed from that run,
which is why the publish workflow does not build it again.

### Then

```bash
git tag -a v<version> -m "v<version>: <one line>"
git push --follow-tags
gh release create v<version> --verify-tag --notes-file <file>   # this is the publish
```

The tag has to exist before the changelog's `[x.y.z]` link resolves, and
`--verify-tag` refuses to invent one that does not. Use the changelog section as
the release body.

Then watch it with `gh run watch`, and check the npm page shows the new version
with a provenance badge. If the job fails after the release exists, re-run it
from the Actions tab; the release does not need cutting again. If it failed
*after* the upload, the version is gone for good, because npm does not reuse a
version number, so the fix is the next patch, not a retry.

## Commits

Conventional-commit subjects (`feat:`, `fix:`, `docs:`, `refactor:`, `chore:`),
scoped where it helps (`feat(parse):`). Write the subject as what the change does
for a reader of the repo, not as a plan step.

`pnpm run check` must be clean; `lefthook` pre-commit jobs run `oxlint --fix` and
then `oxfmt` on staged files if hooks are wired in your environment.
