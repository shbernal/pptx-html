# AGENTS.md

Guidance for coding agents working in the `pptx-html/` project.

The **why** lives in [`docs/`](./docs/index.md) and is not repeated here. The
**how** lives in [`CONTRIBUTING.md`](./CONTRIBUTING.md): setup, commands, test
layers, where to look things up, upstream filing. This file is the working rules
and the routing between them.

## Read before you change

| If you are touching | Read first |
| --- | --- |
| Anything affecting fidelity, the oracle, the corpus | [docs/round-trip.md](./docs/round-trip.md) |
| `src/import`, `src/render`, `src/parse`, `src/emit`, `src/ir` | [docs/architecture.md](./docs/architecture.md) |
| Anything that looks like a missing feature | [docs/decisions.md](./docs/decisions.md) |
| Build, tests, verification, upstream asks | [CONTRIBUTING.md](./CONTRIBUTING.md) |
| `site/**`, the public site | [CONTRIBUTING.md](./CONTRIBUTING.md#the-site) |

## Repository expectations

- This repo builds `pptx-html`: it moves slides between HTML and PPTX by driving
  `@shbernal/ts-pptx`. Its role is the **HTML ⇄ ts-pptx** link, both directions.
- This is a **standalone repository** with its own git. It is not a member of any
  workspace; run its commands from the repository root.
- Use `pnpm`. Node `>=24`. Keep source in `src/`, tests in `test/`. Treat `dist/`
  as generated build output.
- **`@shbernal/ts-pptx` is a released npm range, no longer a git sha.** Bumping it
  is three commands, not one; the skill reinstall is part of it →
  [CONTRIBUTING](./CONTRIBUTING.md#setup).
- **Version numbers live in `package.json`, never in prose.** A range written down
  twice goes stale in one of the two places, and `@shbernal/ts-pptx` moves fast
  enough that it did: the README named a range the manifest had already left four
  minor versions behind. Markdown names the package and points at the manifest.
  The exception is history, where the number *is* the fact: the changelog, and a
  sentence recording which release a fix landed in.
- **Skills live in `.agents/skills/`**, runtime agnostic, with each runtime's own
  directory (`.claude/skills/`, `.gitignore`d) linking into it, so every runtime
  loads the same files and there is one copy to edit.
- **`ts-pptx-upstream` is installed from the dependency, not authored here.** It
  ships inside `@shbernal/ts-pptx`, which is what keeps it matching the installed
  version; editing the copy would only diverge from it, and a change to it belongs
  in the ts-pptx repo. `skills-lock.json` is the tracked record and the copy is
  `.gitignore`d. Install it after a fresh clone, and again after bumping the
  dependency:

  ```bash
  npx skills add ./node_modules/@shbernal/ts-pptx -s '*' -a claude-code -a codex -a universal -y
  ```

  Name the runtimes rather than passing `--all`: that flag writes an `agent/`
  directory at the repo root for a runtime nobody here uses.
  `skills experimental_install` is not the restore command: it repopulates
  `.agents/skills/` from the lock file but creates none of the runtime links.

  Skills this repo does write are repo-scoped only, and stay tracked. Anything
  not specific to this project belongs in the personal skills repo instead.
- Preserve unrelated dirty state. Do not revert user changes.

## The non-negotiables

Each of these has a reason recorded in `docs/`; none of them is a style
preference. If one seems wrong, read the rationale before changing it.

- **No rasterizer in `src/`, and no slide images on the site.** A flattened slide
  is the one output that can never re-enter the loop, and a screenshot of a
  converted deck is that same impression sold to someone who cannot check it.
  Previews are rendered by `renderDeck` at runtime or they are not shown. →
  [decisions](./docs/decisions.md)
- **No raw OOXML, and no post-write repair layer.** The last one deleted every
  speaker note in the deck and nobody noticed from outside. →
  [decisions](./docs/decisions.md)
- **Nothing in `src/emit/` may import `ir/render`.** Emitting from the paint model
  means two writers, and the oracle is only judging one of them. →
  [architecture](./docs/architecture.md)
- **Do not merge the two lanes.** They have different contracts; merging weakens
  the strong one without any test going red. `src/heuristic/model.ts` is not
  `RenderIr` and must not become it. → [architecture](./docs/architecture.md)
- **Do not coin a local `modeled`/`carried`/`unsupported` enum.** The
  machine-readable classification is upstream's `FidelityNote`. Two vocabularies
  for one concept is how the differ and the renderer drift apart. →
  [architecture](./docs/architecture.md)
- **Do not gate on anything but `undeclared`.** `declared` is the contract
  working, `added` is the write path being explicit, `unmatchedNotes` are usually
  invisible to the read model. Never "fix" one by deleting a note. →
  [round-trip](./docs/round-trip.md)
- **Do not add deterministic-looking guarantees to the docs beyond what the oracle
  actually gates.** Fidelity claims stay honest. This binds hardest on `site/**`,
  which is read by people who cannot run the suite: every page touching fidelity
  names the input domain and says PowerPoint-authored decks are not gated.
- **A fidelity change is not done until the oracle covers it.**

## Verification

Run what the change touches. The table is in
[CONTRIBUTING.md](./CONTRIBUTING.md#what-to-run-for-which-change). The two that
are easy to skip and expensive to skip:

- `pnpm run test:oracle` for anything touching fidelity, and update the coverage
  snapshot deliberately, never with a blind `-u`.
- Read `dist/index.d.ts` after changing what `src/index.ts` exports. Nothing
  in-repo imports through the entry point, so a broken public surface is invisible
  to `tsc` and to every test.

Two habits that have each paid for themselves more than once:

- **Probe the write leg; do not read it off the types.** What the writer emits for
  a given option is not in its type signature and is in its source only if you find
  the right function. A throwaway script in `.tmp/` that builds a deck and prints
  the XML answers the question in one run, and twice now that script has become the
  upstream reproduction unchanged.
- **When a property joins the editable surface, check the test that used it as the
  out-of-surface example.** Such a test keeps passing while proving the opposite of
  what it was written to prove. Move it to a property that is still out, and add
  the mirror case.

## Upstream

Gaps in `@shbernal/ts-pptx` are filed as GitHub issues **in the same unit of work
that found them**, before the commit. The `ts-pptx-upstream` skill, which the
package itself ships, is the normative reference for *how*: which error class
means whose bug, reducing the failure to a script that builds its own deck rather
than attaching a real one, which form to file under.

Two things it cannot know, because they are this repo's:

- **What not to file.** Anything about *HTML* stays here: the renderer, the IR
  island, the parser, the extraction heuristics. So does an OOXML limitation,
  which no converter change fixes. The test is whether a fix would help *any*
  consumer of the writer.
- **The triage function.** `FidelityNote.cause`, which `readModelToIr` returns for
  free: `unread` is a missing reader and `unwritable` a missing write option, both
  worth filing; `unsupported` is the output tier's own limit and is not.

There is deliberately no local post-write repair layer, so an issue upstream *is*
the remedy. The half-cycle after the fix is a release rather than a commit, now
that the dependency is a published range: watch for one with `gh release list
--repo shbernal/ts-pptx`, bump the range, delete the workarounds their comments
point at, and close each issue with the test that now passes rather than with
"done". A `github:` install is for *testing* an unreleased fix, not for shipping
against one. This package is published, and a git URL in `dependencies` travels
to every consumer.

Anything discovered while building the custGeom/SVG-path vectorizer goes upstream,
not into a local patch: a missing custGeom case, a measure gap.
