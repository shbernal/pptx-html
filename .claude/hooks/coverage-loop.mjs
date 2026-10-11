#!/usr/bin/env node
/**
 * The coverage loop's guard rails, as Claude Code hooks.
 *
 *   node .claude/hooks/coverage-loop.mjs guard   PreToolUse: keep the agent off the harness
 *   node .claude/hooks/coverage-loop.mjs stop    Stop: no ending a turn on red, dirty or unrecorded
 *
 * Both are inert unless the session was started with `COVERAGE_LOOP=1`, so an
 * ordinary session in this repo can edit the oracle like any other file. See the
 * coverage-loop skill for why each rule exists.
 */

import { execFileSync, spawnSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { relative } from 'node:path'

if (process.env.COVERAGE_LOOP !== '1') process.exit(0)

const input = JSON.parse(readFileSync(0, 'utf8'))
const root = process.env.CLAUDE_PROJECT_DIR ?? input.cwd ?? process.cwd()

/**
 * What the loop is judged by. Editing any of it changes the judge, not the score,
 * so it is the human's to change.
 */
const PROTECTED = [
	/^test\/oracle\/(roundtrip|coverage|surface)(\.test)?\.ts$/,
	/^test\/oracle\/(html|script)-lane\.ts$/,
	/^test\/oracle\/__snapshots__\//,
	/^test\/surface\/(facts|probe|inventory|score)\.ts$/,
	/^test\/surface\/ledger\.json$/,
	/^vitest\.config\.ts$/,
	/^\.claude\//,
	/^\.agents\/skills\/coverage-loop\//,
	/^AGENTS\.md$/,
]

function isProtected(path) {
	const rel = relative(root, path.startsWith('/') ? path : `${root}/${path}`)
	return PROTECTED.some((pattern) => pattern.test(rel))
}

function deny(reason) {
	console.log(
		JSON.stringify({
			hookSpecificOutput: { hookEventName: 'PreToolUse', permissionDecision: 'deny', permissionDecisionReason: reason },
		})
	)
	process.exit(0)
}

function block(reason) {
	console.log(JSON.stringify({ decision: 'block', reason }))
	process.exit(0)
}

const mode = process.argv[2]

if (mode === 'guard') {
	const tool = input.tool_name
	const args = input.tool_input ?? {}
	if (['Edit', 'Write', 'MultiEdit', 'NotebookEdit'].includes(tool)) {
		const path = args.file_path ?? args.notebook_path
		if (path && isProtected(path))
			deny(
				`${path} is part of the oracle the coverage loop is judged by. A probe that needs it changed is a finding: stop and report it rather than editing the judge.`
			)
	}
	if (tool === 'Bash') {
		const command = String(args.command ?? '')
		const rules = [
			[/\bvitest\b[^|;&]*\s(-u|--update)\b/, 'No snapshot updates in the loop: the coverage snapshot changes deliberately, by a human.'],
			[/--no-verify\b/, 'Commit through the hooks.'],
			[/\bgit\s+(push|rebase|reset\s+--hard|commit\b[^|;&]*--amend)\b/, 'History stays one item per commit, local, and append-only during a run.'],
		]
		for (const [pattern, reason] of rules) if (pattern.test(command)) deny(reason)
		const mutates = /(^|[\s;&|])(sed\s+-i|rm|mv|cp|tee|truncate|perl|python3?|node\s+-e)\b|>/.test(command)
		const touched = command.split(/[\s'"=]+/).filter((token) => /[./]/.test(token) && isProtected(token))
		if (mutates && touched.length > 0) deny(`${touched.join(', ')} is part of the oracle the coverage loop is judged by.`)
	}
	process.exit(0)
}

if (mode === 'stop') {
	const dirty = execFileSync('git', ['status', '--porcelain'], { cwd: root, encoding: 'utf8' }).trim()
	if (dirty)
		block(
			`Uncommitted changes:\n${dirty}\nFinish the item (score green, ledger accepted, one commit) or revert it before stopping.`
		)

	const checks = spawnSync('sh', ['-c', 'pnpm run -s build >/dev/null && pnpm run -s typecheck && pnpm run -s check'], {
		cwd: root,
		encoding: 'utf8',
	})
	if (checks.status !== 0)
		block(`build/typecheck/check failed:\n${`${checks.stdout}${checks.stderr}`.trim().split('\n').slice(-30).join('\n')}`)

	const score = spawnSync('node', ['test/surface/score.ts', '--json', '--frontier', '0'], { cwd: root, encoding: 'utf8' })
	let result
	try {
		result = JSON.parse(score.stdout)
	} catch {
		block(`oracle:score did not produce a result:\n${score.stderr.trim().split('\n').slice(-30).join('\n')}`)
	}
	if (!result.ok) {
		const reasons = []
		if (!result.gate) reasons.push('the oracle is red: run `pnpm run test:oracle` and fix it')
		for (const line of result.regressions) reasons.push(`regression: ${line}`)
		if (result.improvements.length > 0)
			reasons.push(`${result.improvements.length} probe change(s) not recorded: run \`pnpm run oracle:score --accept\` and commit the ledger`)
		block(reasons.join('\n'))
	}
	process.exit(0)
}

console.error(`unknown mode ${mode}`)
process.exit(2)
