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
 *    coverage run works from, and its render frontier: the probes whose option
 *    changed the XML but never reached the paint model, with neither `invisible`
 *    nor `unread` to say why.
 *
 * Coverage comes in two columns. XML coverage counts option paths a probe or
 * `SKIPPED` accounts for. Render coverage counts probes, since the paint check is
 * per probe: those that rendered, or that say why not, over all that wrote
 * something (an inert or crashed probe has nothing to paint).
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
	paint: 'rendered' | 'invisible' | 'unread' | 'unrendered' | 'moot'
}

const ROOT = join(import.meta.dirname, '..', '..')
const LEDGER = join(ROOT, 'test', 'surface', 'ledger.json')
const OBSERVATIONS = join(ROOT, '.tmp', 'surface', 'observations.json')

/**
 * Higher is better. A probe may only move up without a human editing the ledger.
 * The round trip decides the rank. Within a status, the lane decides: `survives/template`
 * ranks lowest of the survivals, since the HTML never held it, so relabelling an entry to
 * it is a weaker claim a human has to accept. The render frontier breaks the remaining
 * ties the same way the XML one does: rendered beats accounted for (`invisible`,
 * `unread`), which beats `unrendered`. So a probe that stops reaching the renderer is a
 * regression even while its XML still comes back, and so is dropping the annotation
 * that took it off the frontier.
 */
function rank(state: string): number {
	const [status, lane, paint = ''] = state.split('/')
	const painted = { '': 2, invisible: 1, unread: 1 }[paint] ?? 0
	if (status === 'survives') return (lane === 'authored' ? 4 : lane === 'template' ? 2 : 3) * 3 + painted
	if (status === 'upstream' || status === 'local') return 3 + painted
	return status === 'crash' ? 3 : 0
}

/** A probe's ledger state: `status/lane`, plus its paint state unless it rendered or had nothing to render. */
const stateOf = (seen: ObservationRecord) =>
	`${seen.status}/${seen.lane}${seen.paint === 'rendered' || seen.paint === 'moot' ? '' : `/${seen.paint}`}`

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
	const paints: Record<string, number> = {}
	for (const seen of probes) {
		const key = seen.status === 'survives' ? `survives/${seen.lane}` : seen.status
		statuses[key] = (statuses[key] ?? 0) + 1
		if (seen.paint !== 'moot') paints[seen.paint] = (paints[seen.paint] ?? 0) + 1
	}
	const paintable = probes.filter((seen) => seen.paint !== 'moot')
	return {
		family,
		paths: own.length,
		probed: own.filter((option) => covered.has(option.path)).length,
		skipped: own.filter((option) => skippedSet.has(option.path)).length,
		todo: own.filter((option) => !covered.has(option.path) && !skippedSet.has(option.path)),
		statuses,
		render: { probes: paintable.length, done: paintable.filter((seen) => seen.paint !== 'unrendered').length },
		paints,
		renderTodo: paintable.filter((seen) => seen.paint === 'unrendered').map((seen) => seen.id),
	}
})

// 3. The ratchet.
const ledger: Record<string, string> = existsSync(LEDGER) ? JSON.parse(readFileSync(LEDGER, 'utf8')) : {}
const current: Record<string, string> = Object.fromEntries(
	observations.map((seen) => [seen.id, stateOf(seen)] as const).toSorted(([a], [b]) => a.localeCompare(b))
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
// Counted over observations, not family rows: a probe whose covers span two
// families shows in both rows and must count once here.
const paintable = observations.filter((seen) => seen.paint !== 'moot')
const renderTotals = {
	probes: paintable.length,
	done: paintable.filter((seen) => seen.paint !== 'unrendered').length,
}
const scope = args.family ? families.filter((row) => row.family === args.family) : families
const frontier = scope.flatMap((row) => row.todo)
const renderFrontier = [...new Set(scope.flatMap((row) => row.renderTodo))]
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
				render: renderTotals,
				families: families.map(({ todo, renderTodo, ...row }) => ({
					...row,
					todo: todo.length,
					renderTodo: renderTodo.length,
				})),
				frontier: frontier.slice(0, Number(args.frontier)),
				renderFrontier: renderFrontier.slice(0, Number(args.frontier)),
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
	console.log(
		`render      ${renderTotals.done}/${renderTotals.probes} probes rendered or accounted for (${pct(renderTotals.done, renderTotals.probes)})`
	)
	console.log('')
	console.log('family               paths  probed  skipped  todo  render  rtodo   probe results | paint')
	const counts = (record: Record<string, number>) =>
		Object.entries(record)
			.map(([key, count]) => `${key}=${count}`)
			.join(' ')
	for (const row of families) {
		console.log(
			`${row.family.padEnd(20)} ${String(row.paths).padStart(5)}  ${String(row.probed).padStart(6)}  ${String(row.skipped).padStart(7)}  ${String(row.todo.length).padStart(4)}  ${pct(row.render.done, row.render.probes).padStart(6)}  ${String(row.renderTodo.length).padStart(5)}   ${counts(row.statuses)} | ${counts(row.paints)}`
		)
	}
	for (const line of regressions) console.log(`REGRESSION  ${line}`)
	if (!accepted) for (const line of improvements) console.log(`changed     ${line}`)
	if (args.family) {
		console.log('')
		console.log(`frontier (${args.family}, next ${args.frontier} of ${frontier.length}):`)
		for (const option of frontier.slice(0, Number(args.frontier))) console.log(`  ${option.path}  ::  ${option.type}`)
		console.log('')
		console.log(`render frontier (${args.family}, next ${args.frontier} of ${renderFrontier.length}):`)
		for (const id of renderFrontier.slice(0, Number(args.frontier))) console.log(`  ${id}`)
	}
	if (!gate) console.log(`\n--- oracle output (tail) ---\n${gateTail}`)
}

process.exit(ok ? 0 : 1)
