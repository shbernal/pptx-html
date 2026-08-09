---
name: ts-pptx-upstream
description: File gaps in @shbernal/ts-pptx as GitHub issues on shbernal/ts-pptx the moment they surface, instead of noting them in chat or in a doc. Use when work in pptx-html hits a missing read accessor, a write option that cannot express the source, a writer bug that would otherwise need a post-write workaround, a FidelityNote with cause `unread` or `unwritable`, or an oracle lane that loses something on a round trip — and also when checking whether an ask is already filed, when a new ts-pptx release lands and workarounds may be deletable, or when deciding whether a problem belongs upstream at all. Not for problems in this repo's own HTML side (renderer, island, parser, extraction heuristics) and not for OOXML limitations, which no upstream change can fix.
---

# Filing upstream asks against `@shbernal/ts-pptx`

The dependency is pinned at `^3.0.0` and consumed from npm — no local link, no
workspace override, no vendored patch. The only channel to upstream is a GitHub
issue on **`shbernal/ts-pptx`** (public, `gh` authenticated on this machine).

## File on arrival

**A gap that is only mentioned in a chat session or written into a doc does not
exist.** Sessions end and plan files are deleted by design; the issue tracker is
the one queue that survives both. So when a gap surfaces, file it in the same
unit of work that found it — before the commit, not batched into a later sweep.

An ask that turns out to be unclear or misdirected gets closed, and that is a
cheap outcome. A gap that was never filed costs a rediscovery from scratch,
usually by someone who then re-invents the same workaround.

Two things go with the filing, and neither is optional:

- The workaround in the code carries a comment naming the **issue URL** and the
  condition under which it is deleted.
- The ask is recorded where the project tracks blocked work, with **issue → what
  it blocks → the release that carries it**.

## What to file, and what not to

`FidelityNote.cause` from `@shbernal/ts-pptx/script` is the triage function and
it comes back from `readModelToIr` for free:

| `cause` | Meaning | File? |
|---|---|---|
| `unread` | the read API exposes no accessor for something in the deck | **yes** — a missing reader |
| `unwritable` | reads fine, the write API cannot express it | **yes** — a missing writer option |
| `unsupported` | OOXML or output-tier limitation | **no** — no converter work fixes it |

Also file: any writer bug this repo would otherwise have to work around after the
fact, and any round-trip loss an oracle lane has to declare for itself because
upstream's own note set does not cover the case. There is deliberately no local
post-write repair layer — the one that existed was deleted for hiding both its
own obsolescence and its own damage — so an issue here *is* the remedy.

Do **not** file: anything about HTML — the renderer, the IR island, the parser,
the extraction heuristics. The round-trip guarantee itself is a property of this
loop, not of the writer. The rule of thumb is whether a fix would help *any*
consumer of the writer.

Before filing, check the ask is real against the installed typings rather than
from memory — `grep -rn '<accessor>' node_modules/@shbernal/ts-pptx/dist/*.d.ts`.
A missing accessor whose twin exists on the write side is the strongest form of
the ask, because the asymmetry is the argument.

## How to file

Write the body to a file and pass `--body-file`. Same reason commit messages go
through a file: this machine has both a POSIX shell and PowerShell, they
disagree about here-doc syntax, and a wrong dialect is passed through as literal
text instead of raising an error.

```bash
gh issue create --repo shbernal/ts-pptx \
  --title "[FEATURE] read: <what is missing>" \
  --label enhancement \
  --body-file <scratchpad>/ask.md
```

Titles take the repo template's `[BUG]` / `[FEATURE]` prefix. Labels are the
stock GitHub set — `enhancement` for a missing accessor or option, `bug` for
wrong or under-reported output. Do not invent a label taxonomy in that repo.

The body follows `.github/ISSUE_TEMPLATE/user-template.md` there: issue
category, product versions (the installed `@shbernal/ts-pptx`, Node, pnpm),
**Desired Behavior**, **Observed Behavior**, **Steps to Reproduce**. Fetch it if
in doubt:

```bash
gh api repos/shbernal/ts-pptx/contents/.github/ISSUE_TEMPLATE/user-template.md --jq .content | base64 -d
```

What makes an ask actionable months later, in this order:

1. The **XML** it is about (`spPr/a:noFill`, `a:buAutoNum/@startAt`) — the one
   detail that never goes stale.
2. The **shape of the API** being asked for, named concretely, and its existing
   twin if there is one.
3. What a consumer is forced to do without it, and why that is wrong rather than
   merely inconvenient. "The model reports something the deck did not say" is a
   much stronger case than "I would like this".
4. The workaround in place here and the condition for deleting it.

## After a release lands

A filed ask is not done until the local stopgap is gone. When a new version is
published: `gh issue list --repo shbernal/ts-pptx --state all`, check which asks
the release carries, bump the pin, delete the workarounds their comments point
at, and close each issue with the evidence — the test that now passes, not
"done".
