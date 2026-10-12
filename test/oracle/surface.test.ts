import { mkdirSync, writeFileSync } from 'node:fs'
import { afterAll, describe, expect, it } from 'vitest'
import { inventory } from '../surface/inventory'
import { disagreement, type Observation, observe } from '../surface/probe'
import { PROBES, SKIPPED } from '../surface/probes'

// The surface probes: each writer option the manifest names, through the whole
// loop, judged on the XML rather than on the read model. See `test/surface/`.
//
// Every observation is also written to `.tmp/surface/observations.json`, which
// `pnpm run oracle:score` turns into the score and the frontier. The file is an
// output of this run, never an input to it.

const PATHS = new Set(inventory().map((option) => option.path))
const observations: Observation[] = []

describe('surface probes', () => {
	it('the inventory still finds the options the corpus is known to use', () => {
		// The reader is structural, so a change in how the writer's declarations are
		// emitted could empty it without an error. These are options the corpus sets.
		for (const path of [
			'addText.bold',
			'addText.underline.style',
			'addShape.line.dashType',
			'addTable.cell.fill.color',
		])
			expect(PATHS, path).toContain(path)
		expect(PATHS.size).toBeGreaterThan(500)
	})

	it('probe ids are unique: the ledger keys on them', () => {
		expect(new Set(PROBES.map((probe) => probe.id)).size).toBe(PROBES.length)
	})

	it('every covered and skipped path is a real option', () => {
		const unknown = [...PROBES.flatMap((probe) => probe.covers), ...Object.keys(SKIPPED)].filter(
			(path) => !PATHS.has(path)
		)
		expect(unknown).toEqual([])
	})

	it('no path is both probed and skipped', () => {
		const covered = new Set(PROBES.flatMap((probe) => probe.covers))
		expect(Object.keys(SKIPPED).filter((path) => covered.has(path))).toEqual([])
	})

	for (const probe of PROBES) {
		it(`${probe.id}`, async () => {
			const seen = await observe(probe)
			observations.push(seen)
			const problem = disagreement(probe, seen)
			const detail = [
				`status ${seen.status} (${seen.lane}${seen.rendered ? '' : ', not rendered'}), oracle: ${seen.undeclared} undeclared / ${seen.declared} declared`,
				...seen.htmlLost.slice(0, 12).map((fact) => `  html lost   ${fact}`),
				...seen.scriptLost.slice(0, 12).map((fact) => `  script lost ${fact}`),
			].join('\n')
			expect(problem, `${probe.id}: ${problem}\n${detail}`).toBeNull()
		})
	}

	afterAll(() => {
		mkdirSync('.tmp/surface', { recursive: true })
		writeFileSync(
			'.tmp/surface/observations.json',
			`${JSON.stringify({ observations, skipped: Object.keys(SKIPPED) }, null, '\t')}\n`
		)
	})
})
