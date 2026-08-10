/**
 * charcheck, pointed at the website.
 *
 * The scope is the site and the prose it renders, not the library: nothing under `src/`
 * reaches a reader of the site, and `test/` even less so.
 *
 * Everything here is `error`. It was `warn` for exactly as long as it took to read the
 * 129 findings and reword them: a gate is only honest once the surface behind it is
 * clean, and a warning nobody has to act on is a warning everybody scrolls past.
 *
 * So this now blocks rather than reports. The commit that cleared the backlog is the
 * same one that closed the gate, which is the only ordering that leaves no window where
 * a dash could be reintroduced unnoticed.
 *
 * The characters are written as escapes rather than literally, so this file is not itself
 * a finding in the rule it defines.
 */

import { strategies } from 'charcheck/config'

/**
 * Em dash and the horizontal bar, with the whitespace on either side. That whitespace is
 * part of the match so `--fix` replaces `a — b` with `a: b` rather than leaving `a:  b`.
 *
 * `\s` rather than `[ \t]`, so a dash sitting at the end or the start of a wrapped line
 * takes the line break with it and the two lines join. Restricting it to horizontal space
 * was tried first and is worse: it leaves a trailing space behind a line-final dash, and
 * eats the indent in front of a line-initial one. The join re-flows the paragraph, which
 * needs a re-wrap, which is a thing you can see in the diff you were told to read.
 */
const CLAUSE_DASH = '\\s*[\\u2014\\u2015]\\s*'

const REWORD = 'Reads as an em dash on the page. A colon, a comma, or reword.'

/** Shared by every rule below. `fix` makes `--fix` propose a replacement to read over. */
const dashRule = (id, include, scope) => ({
	id,
	pattern: CLAUSE_DASH,
	fix: strategies.clauseSeparator,
	severity: 'error',
	message: REWORD,
	include,
	...(scope ? { scope } : {}),
})

export default {
	rules: [
		// Pages, and the docs the pages are generated from. `docs/` is the tracked source;
		// `site/docs/` is a mirror written by site/scripts/sync-docs.ts, so scanning the
		// mirror would report every finding twice and point at a file nobody edits.
		dashRule('no-em-dash-in-site-prose', ['docs/**/*.md', 'site/*.md']),

		// Theme sources: deck copy, slide labels, UI strings. Only text that can reach the
		// page, so a dash in a comment here stays allowed. Note the pattern names
		// `.vitepress` explicitly: a dotted directory is only entered when one does.
		dashRule('no-em-dash-in-site-strings', ['site/.vitepress/**/*.ts'], 'strings'),

		// Build scripts, which hold two kinds of string this rule cannot tell apart: page
		// prose that build-ledger.ts writes into the Evidence page, and error messages
		// that only ever reach whoever broke the build. The scope is token-based, not an
		// AST, so a `throw` argument looks exactly like a heading. The three error
		// messages in docs-source.ts carry a `charcheck-disable-next-line` naming this
		// rule; anything else it reports is prose a reader sees.
		dashRule('no-em-dash-in-site-build-strings', ['site/scripts/**/*.ts'], 'strings'),

		// Components: template text, allowlisted attributes and script literals. `<style>`
		// blocks and template comments are exempt.
		dashRule('no-em-dash-in-site-markup', ['site/.vitepress/**/*.vue'], 'markup'),

		// The commit message, reached only when charcheck is given `--commit-msg`.
		//
		// `error` from the start, unlike every rule above, because this surface has no
		// backlog to grandfather: each message is text nobody has written yet. The cost of
		// a wrong one is a `git commit --amend`, not 129 rewordings.
		//
		// No `fix`. `--fix` is refused for a commit message on purpose, so declaring one
		// would only advertise a repair that cannot run.
		{
			id: 'no-em-dash-in-commit-msg',
			pattern: CLAUSE_DASH,
			severity: 'error',
			message: 'Reads as an em dash in the log. A colon, a comma, or reword.',
			include: ['<commit-msg>'],
		},
	],

	ignore: ['site/docs/**', 'site/.vitepress/cache/**', 'site/.vitepress/dist/**'],
}
