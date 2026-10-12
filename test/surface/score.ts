/**
 * `pnpm run oracle:score`: the gate, the surface score, the ratchet and the frontier.
 *
 * ```
 * node test/surface/score.ts [--family addText] [--accept] [--json]
 * ```
 *
 * 1. **Gate.** Runs the whole oracle project. Red is red; nothing below excuses it.
 * 2. **Score.** Reads the observations the surface probes just wrote and counts,
 *    per family, how many inventory paths a probe or `SKIPPED` accounts for, and
 *    how the probes came out.
 * 3. **Ratchet.** Compares every probe against `ledger.json`, the last accepted
 *    state. A probe that disappeared or got worse is a regression and fails the
 *    run. `--accept` records the current state, and refuses while anything is red.
 * 4. **Frontier.** The next unaccounted-for paths of `--family`, which is what a
 *    coverage run works from.
 *
 * Exit status is 0 only when the gate is green, nothing regressed, and the ledger
 * is current. That last one is deliberate: a probe nobody accepted is progress
 * nobody recorded, and the loop's stop hook will not let a run end on it.
 *
 * Plain `node`, no vitest: it imports nothing but `inventory.ts`, and it learns
 * everything else from the observations file the test run writes.
 */

import { spawnSync } from 'node:child_process'
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { parseArgs } from 'node:util'
import { FAMILIES, inventory } from './inventory.ts'

interface ObservationRecord {
	id: string
	covers: string[]
	status: 'survives' | 'upstream' | 'local' | 'inert' | 'crash'
	lane: 'authored' | 'carried' | 'template' | 'mixed'
}

const ROOT = join(import.meta.dirname, '..', '..')
const LEDGER = join(ROOT, 'test', 'surface', 'ledger.json')
const OBSERVATIONS = join(ROOT, '.tmp', 'surface', 'observations.json')

/**
 * Higher is better. A probe may only move up without a human editing the ledger.
 * `survives/template` ranks lowest of the survivals: the HTML never held it, so
 * relabelling an entry to it is a weaker claim, which a human has to accept.
 */
function rank(state: string): number {
	const [status, lane] = state.split('/')
	if (status === 'survives') return lane === 'authored' ? 4 : lane === 'template' ? 2 : 3
	if (status === 'upstream' || status === 'local' || status === 'crash') return 1
	return 0
}

const { values: args } = parseArgs({
	options: {
		family: { type: 'string' },
		accept: { type: 'boolean', default: false },
		json: { type: 'boolean', default: false },
		frontier: { type: 'string', default: '25' },
	},
})

if (args.family && !FAMILIES.includes(args.family)) {
	console.error(`unknown family ${args.family}; one of: ${FAMILIES.join(', ')}`)
	process.exit(2)
}

// 1. The gate: the whole oracle project, surface probes included.
const run = spawnSync('pnpm', ['exec', 'vitest', 'run', '--project', 'oracle'], { cwd: ROOT, encoding: 'utf8' })
const gate = run.status === 0
const gateTail = `${run.stdout}${run.stderr}`.trim().split('\n').slice(-60).join('\n')

if (!existsSync(OBSERVATIONS)) {
	console.error(`no observations at ${OBSERVATIONS}; the oracle run did not reach the surface probes\n${gateTail}`)
	process.exit(1)
}
const { observations, skipped } = JSON.parse(readFileSync(OBSERVATIONS, 'utf8')) as {
	observations: ObservationRecord[]
	skipped: string[]
}

// 2. The score.
const paths = inventory()
const covered = new Set(observations.flatMap((seen) => seen.covers))
const skippedSet = new Set(skipped)
const families = FAMILIES.map((family) => {
	const own = paths.filter((option) => option.family === family)
	const probes = observations.filter((seen) => seen.covers.some((path) => own.some((option) => option.path === path)))
	const statuses: Record<string, number> = {}
	for (const seen of probes) {
		const key = seen.status === 'survives' ? `survives/${seen.lane}` : seen.status
		statuses[key] = (statuses[key] ?? 0) + 1
	}
	return {
		family,
		paths: own.length,
		probed: own.filter((option) => covered.has(option.path)).length,
		skipped: own.filter((option) => skippedSet.has(option.path)).length,
		todo: own.filter((option) => !covered.has(option.path) && !skippedSet.has(option.path)),
		statuses,
	}
})

// 3. The ratchet.
const ledger: Record<string, string> = existsSync(LEDGER) ? JSON.parse(readFileSync(LEDGER, 'utf8')) : {}
const current: Record<string, string> = Object.fromEntries(
	observations
		.map((seen) => [seen.id, `${seen.status}/${seen.lane}`] as const)
		.toSorted(([a], [b]) => a.localeCompare(b))
)
const regressions: string[] = []
const improvements: string[] = []
for (const [id, was] of Object.entries(ledger)) {
	const now = current[id]
	if (now === undefined) regressions.push(`${id}: was ${was}, probe is gone`)
	else if (rank(now) < rank(was)) regressions.push(`${id}: was ${was}, now ${now}`)
	else if (now !== was) improvements.push(`${id}: ${was} → ${now}`)
}
for (const id of Object.keys(current)) if (!(id in ledger)) improvements.push(`${id}: new, ${current[id]}`)

let accepted = false
if (args.accept) {
	if (!gate || regressions.length > 0) {
		console.error('refusing --accept: the gate is red or something regressed')
	} else {
		writeFileSync(LEDGER, `${JSON.stringify(current, null, '\t')}\n`)
		accepted = true
	}
}
const ledgerCurrent = accepted || improvements.length === 0

const totals = families.reduce(
	(sum, row) => ({ paths: sum.paths + row.paths, done: sum.done + row.probed + row.skipped }),
	{ paths: 0, done: 0 }
)
const frontier = (args.family ? families.filter((row) => row.family === args.family) : families).flatMap(
	(row) => row.todo
)
const ok = gate && regressions.length === 0 && ledgerCurrent

if (args.json) {
	console.log(
		JSON.stringify(
			{
				ok,
				gate,
				regressions,
				improvements: accepted ? [] : improvements,
				coverage: totals,
				families: families.map(({ todo, ...row }) => ({ ...row, todo: todo.length })),
				frontier: frontier.slice(0, Number(args.frontier)),
			},
			null,
			'\t'
		)
	)
} else {
	const pct = (n: number, d: number) => (d === 0 ? '—' : `${((100 * n) / d).toFixed(1)}%`)
	console.log(`gate        ${gate ? 'pass' : 'FAIL'}`)
	console.log(`ratchet     ${regressions.length === 0 ? 'ok' : `${regressions.length} REGRESSION(S)`}`)
	console.log(
		`ledger      ${accepted ? 'accepted' : ledgerCurrent ? 'current' : `${improvements.length} change(s) not accepted, run with --accept`}`
	)
	console.log(`coverage    ${totals.done}/${totals.paths} option paths (${pct(totals.done, totals.paths)})`)
	console.log('')
	console.log('family               paths  probed  skipped  todo   probe results')
	for (const row of families) {
		const statuses = Object.entries(row.statuses)
			.map(([key, count]) => `${key}=${count}`)
			.join(' ')
		console.log(
			`${row.family.padEnd(20)} ${String(row.paths).padStart(5)}  ${String(row.probed).padStart(6)}  ${String(row.skipped).padStart(7)}  ${String(row.todo.length).padStart(4)}   ${statuses}`
		)
	}
	for (const line of regressions) console.log(`REGRESSION  ${line}`)
	if (!accepted) for (const line of improvements) console.log(`changed     ${line}`)
	if (args.family) {
		console.log('')
		console.log(`frontier (${args.family}, next ${args.frontier} of ${frontier.length}):`)
		for (const option of frontier.slice(0, Number(args.frontier))) console.log(`  ${option.path}  ::  ${option.type}`)
	}
	if (!gate) console.log(`\n--- oracle output (tail) ---\n${gateTail}`)
}

process.exit(ok ? 0 : 1)
