---
name: coverage-loop
description: Work through the pptx-ts option surface one probe at a time, using `pnpm run oracle:score` as the feedback signal. Use when running or resuming a coverage run (usually under `/goal` with COVERAGE_LOOP=1), when adding a surface probe, or when an `oracle:score` frontier needs working down.
---

# The coverage loop

The goal is to know, for every option `pptx-ts` lets an author set, whether it
survives `pptx → HTML → pptx`. You extend `test/surface/probes.ts` one probe at a
time, and the scoreboard tells you what is left and whether you broke anything.

Read `docs/round-trip.md` ("The surface probes") once before starting. It
explains why the probe judges XML rather than `DeckIr`.

## One iteration

1. **Pick.** Run `pnpm run oracle:score --family <family>` and take the first
   frontier path, or the first few if one probe naturally covers them (e.g.
   `underline.style` and `underline.color`).
2. **Probe.** Add an entry to `PROBES` in `test/surface/probes.ts`, using a helper
   (`text`, `shape`, `table`, `cell`) or a new helper in that file for a new
   family. Use a non-default value. List every path it exercises in `covers`.
3. **Run.** `pnpm exec vitest run --project oracle test/oracle/surface.test.ts -t '<id>'`.
   The failure message gives the status and the lost XML facts.
4. **Classify and act** on the status:

   | status | meaning | what to do |
   | --- | --- | --- |
   | `survives` | the option's XML came back | nothing |
   | `upstream` | the script lane, upstream code end to end, loses it too | reduce it to a pptx-ts-only script in `.tmp/`, file it (below), set `issue` |
   | `local` | only the HTML lane loses it | fix it in `src/` (render, parse, emit); if it cannot be fixed this iteration, set `accepted: { kind: 'local', reason }` |
   | `inert` | the option wrote nothing | try a non-default value; if the writer drops it, file it (below) and set `unwritten`; if it truly writes nothing (authoring-time only), set `accepted: { kind: 'inert', reason }` |
   | `crash` | a lane threw | the error class says whose bug it is (see the upstream skill); fix or file |

   `accepted: { kind: 'equivalent' }` is for an output that says the same thing in
   a different spelling. The reason must say how they are equivalent.
   `unsupported` is for losses OOXML or the HTML tier cannot carry. Use either one
   sparingly: every accepted entry is a claim a reviewer will check.

5. **Skip** a path only when probing it would prove nothing: an input source like
   `data`/`path`, or a pagination knob that writes no XML of its own. Put it in
   `SKIPPED` with a reason. "Hard to probe" is not a reason.
6. **Gate.** `pnpm run oracle:score --family <family> --accept`. It refuses while
   the oracle is red or anything regressed.
7. **Commit** probes, fixes and `test/surface/ledger.json` together, one item per
   commit, with a conventional subject. Put the score change in the body. The
   "before" is the `coverage` line step 1 printed, before the probe existed;
   measuring it now counts the new probe on both sides:

   ```
   test(surface): probe addText.underline.color

   coverage 14/1510 → 15/1510; addText.underline.color survives (authored)
   ```

   A local fix is `fix(render): …` and carries its probe in the same commit.

## Filing upstream

Every `upstream` status, and every `crash` that the error class attributes to
`pptx-ts`, becomes an issue on **`shbernal/pptx-ts`**. Do not ask first: the
maintainer pre-authorized it. Follow the `ts-pptx-upstream` skill for the form,
the reduction and the duplicate search, but pass `--repo shbernal/pptx-ts` (the
skill's examples use the repo's old name). Write the body under `.tmp/`. One
issue per distinct cause: if ten options are lost by the same missing reader,
file one issue listing all of them, and point every probe at it.

When the read model has no accessor for the value at all, say so. That is the
strongest form of gap.

## Never

The hooks block these during a run. They are listed so you know why:

- Edit the oracle harness, the lanes, the probe runner, the inventory, the score
  script, the ledger, snapshots, or this skill. A probe that seems to need any of
  that is a finding: **stop and report it**.
- `vitest -u`, `--no-verify`, `git push`, `rebase`, `reset --hard`, `--amend`.
- Change a probe's value until it happens to pass. A loss is recorded, not avoided.
- Delete or weaken a probe. The ratchet fails on a missing probe.

## When to stop

- The family's frontier is empty and the score is green: summarize and stop.
- The same root cause blocks three probes in a row and is filed: say so and
  stop, since the rest of that area is waiting on one fix.
- A probe needs a harness change: describe the change and stop.

Before stopping, the tree must be committed and `oracle:score` must exit 0. The
Stop hook enforces this.

## The summary at the end

Write a summary covering the family, the coverage before and after, the
`survives`/`upstream`/`local`/`accepted` counts, every issue filed (with URL),
every `accepted` entry with its reason, and anything that needs a human:
harness changes, suspected false positives in the fact diff.
