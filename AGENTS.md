# AGENTS.md

Guidance for coding agents working in the `pptx-html/` project.

The **why** lives in [`docs/`](./docs/index.md) and is not repeated here. The
**how** — setup, commands, test layers, where to look things up, upstream filing
— lives in [`CONTRIBUTING.md`](./CONTRIBUTING.md). This file is the working rules
and the routing between them.

## Read before you change

| If you are touching | Read first |
| --- | --- |
| Anything affecting fidelity, the oracle, the corpus | [docs/round-trip.md](./docs/round-trip.md) |
| `src/import`, `src/render`, `src/parse`, `src/emit`, `src/ir` | [docs/architecture.md](./docs/architecture.md) |
| Anything that looks like a missing feature | [docs/decisions.md](./docs/decisions.md) |
| Build, tests, verification, upstream asks | [CONTRIBUTING.md](./CONTRIBUTING.md) |

## Repository expectations

- This repo builds `pptx-html`: it moves slides between HTML and PPTX by driving
  `@shbernal/ts-pptx`. Its role is the **HTML ⇄ ts-pptx** link, both directions.
- This is a **standalone repository** with its own git. It is not a member of any
  workspace; run its commands from the repository root.
- Use `pnpm`. Node `>=24`. Keep source in `src/`, tests in `test/`. Treat `dist/`
  as generated build output.
- **Skills live in `.agents/skills/`** — the tracked source of truth, runtime
  agnostic. `.claude/skills/` is a junction to it (`.gitignore`d), so Claude Code
  loads the same files every other runtime does and there is one copy to edit.
  Recreate it after a fresh clone:

  ```powershell
  New-Item -ItemType Junction -Path .claude\skills -Target .agents\skills
  ```

  ```bash
  ln -s ../.agents/skills .claude/skills
  ```

  Repo-scoped skills only — anything not specific to this project belongs in the
  personal skills repo instead.
- Preserve unrelated dirty state. Do not revert user changes.

## The non-negotiables

Each of these has a reason recorded in `docs/`; none of them is a style
preference. If one seems wrong, read the rationale before changing it.

- **No rasterizer in `src/`.** A flattened slide is the one output that can never
  re-enter the loop. → [decisions](./docs/decisions.md)
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
  actually gates.** Fidelity claims stay honest.
- **A fidelity change is not done until the oracle covers it.**

## Verification

Run what the change touches — the table is in
[CONTRIBUTING.md](./CONTRIBUTING.md#what-to-run-for-which-change). The two that
are easy to skip and expensive to skip:

- `pnpm run test:oracle` for anything touching fidelity, and update the coverage
  snapshot deliberately, never with a blind `-u`.
- Read `dist/index.d.ts` after changing what `src/index.ts` exports. Nothing
  in-repo imports through the entry point, so a broken public surface is invisible
  to `tsc` and to every test.

## Upstream

Gaps in `@shbernal/ts-pptx` are filed as GitHub issues **in the same unit of work
that found them**, before the commit. `.agents/skills/ts-pptx-upstream/` is the
normative reference: what to file, what not to (anything about *HTML* stays here),
how to write an ask that is still actionable months later, and what to do when a
release lands.

Anything discovered while building the custGeom/SVG-path vectorizer — a missing
custGeom case, a measure gap — goes upstream, not patched locally.
