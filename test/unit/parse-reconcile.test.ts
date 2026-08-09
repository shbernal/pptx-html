/**
 * The lane decision and the fold, on a fixture and with no browser in sight.
 *
 * `reconcile` is pure precisely so this file can exist: the DOM work is all in
 * `parse/surface.ts`, so the part with a guarantee attached — *what an edit is
 * allowed to mean* — is testable without Chromium, and the browser only has to
 * prove that reading the DOM is faithful.
 *
 * Every failure here is silent. A lane that reports `exact` for an edited slide
 * loses the edit and says nothing; one that reports `drifted` for an untouched
 * slide makes the signal worthless by crying wolf on every document; and a fold
 * that treats an absent property as "unchanged" makes clearing a property the one
 * edit that is quietly ignored.
 */

import { describe, expect, it } from 'vitest'
import type { RenderIr } from '../../src/ir/render'
import { project, type SanctionedProjection } from '../../src/ir/surface'
import { reconcile } from '../../src/parse/reconcile'
import type { SurfaceReading } from '../../src/parse/surface'
import { SAMPLE_IR } from '../fixtures/render-ir'

/** A reading that says "nothing changed" — the baseline every case below perturbs. */
function unchanged(ir: RenderIr): SurfaceReading {
	return { projection: structuredClone(project(ir)) as SanctionedProjection, deleted: [], anomalies: [] }
}

/** The first run of the first node that has one, by address. */
function firstRun(reading: SurfaceReading) {
	for (const slide of reading.projection.slides) {
		for (const node of slide.nodes) {
			const run = node.runs[0]
			if (run !== undefined) return run
		}
	}
	throw new Error('the fixture has no runs')
}

describe('the lane a slide takes', () => {
	it('is exact when the surface came back identical', () => {
		const result = reconcile(SAMPLE_IR, unchanged(SAMPLE_IR))
		expect(result.slides.map((slide) => slide.lane)).toEqual(SAMPLE_IR.slides.map(() => 'exact'))
		expect(JSON.stringify(result.ir)).toBe(JSON.stringify(SAMPLE_IR))
	})

	it('is reconciled on the edited slide and exact on every other', () => {
		// The per-slide half of the claim. A deck where one slide was edited and the
		// rest were not must not describe them all the same way, or the report tells
		// a caller nothing it could act on.
		const reading = unchanged(SAMPLE_IR)
		firstRun(reading).text = 'Retyped'

		const result = reconcile(SAMPLE_IR, reading)
		expect(result.slides[0]?.lane).toBe('reconciled')
		expect(result.slides[0]?.edits).toBe(1)
		for (const slide of result.slides.slice(1)) expect(slide.lane).toBe('exact')
	})

	it('is drifted when an anomaly names the slide, even alongside a real edit', () => {
		const reading = unchanged(SAMPLE_IR)
		firstRun(reading).text = 'Retyped'
		reading.anomalies.push('run s1.sp3/0/0 was removed from the document; the model’s text was kept')

		const result = reconcile(SAMPLE_IR, reading)
		expect(result.slides[0]?.lane).toBe('drifted')
		expect(result.slides[0]?.notes).toHaveLength(1)
	})

	it('attaches an anomaly to the deck when it names no slide', () => {
		const reading = unchanged(SAMPLE_IR)
		reading.anomalies.push('something went wrong somewhere')

		const result = reconcile(SAMPLE_IR, reading)
		expect(result.warnings).toEqual(['something went wrong somewhere'])
		expect(result.slides.every((slide) => slide.lane === 'exact')).toBe(true)
	})
})

describe('what the fold writes', () => {
	it('leaves the island untouched so the two can still be compared', () => {
		// `modelHash` covers the island, and `editsBetween` diffs against it. A fold
		// that mutated in place would move the baseline with the change and make
		// every edit invisible — which is exactly the bug this file was written after.
		const before = JSON.stringify(SAMPLE_IR)
		const reading = unchanged(SAMPLE_IR)
		firstRun(reading).text = 'Retyped'

		const result = reconcile(SAMPLE_IR, reading)
		expect(JSON.stringify(SAMPLE_IR)).toBe(before)
		expect(JSON.stringify(result.ir)).not.toBe(before)
	})

	it('clears a property that came back absent rather than keeping it', () => {
		// Absence is how this model spells *inherited*, so "absent means unchanged"
		// would make clearing a bold the one edit the surface silently ignores.
		const reading = unchanged(SAMPLE_IR)
		const bolded = reading.projection.slides
			.flatMap((slide) => slide.nodes)
			.flatMap((node) => node.runs)
			.find((run) => run.props.bold !== undefined)
		if (bolded === undefined) throw new Error('the fixture has no run stating bold')
		delete bolded.props.bold

		const result = reconcile(SAMPLE_IR, reading)
		const after = project(result.ir)
			.slides.flatMap((slide) => slide.nodes)
			.flatMap((node) => node.runs)
			.find((run) => run.node === bolded.node && run.paragraph === bolded.paragraph && run.run === bolded.run)
		expect(after?.props).not.toHaveProperty('bold')
		expect(result.slides[0]?.edits).toBe(1)
	})

	it('removes a deleted node and counts it as one edit', () => {
		const reading = unchanged(SAMPLE_IR)
		const doomed = SAMPLE_IR.slides[0]?.nodes[1]
		const slide = reading.projection.slides[0]
		if (doomed === undefined || slide === undefined) throw new Error('the fixture lost its second node')
		reading.deleted.push(doomed.id)
		slide.nodes = slide.nodes.filter((node) => node.id !== doomed.id)

		const result = reconcile(SAMPLE_IR, reading)
		expect(result.ir.slides[0]?.nodes.map((node) => node.id)).not.toContain(doomed.id)
		expect(result.slides[0]?.edits).toBe(1)
	})

	it('removes a deleted node from inside a group and keeps the group', () => {
		// A group whose children all went is still a node the source deck has. The
		// conservative reading is the correct one: inferring "you meant the group too"
		// is exactly the help this layer refuses to give.
		const group = SAMPLE_IR.slides[0]?.nodes.find((node) => node.kind === 'group')
		if (group?.kind !== 'group') throw new Error('the fixture has no group')
		const child = group.children[0]
		if (child === undefined) throw new Error('the fixture group has no children')

		const reading = unchanged(SAMPLE_IR)
		reading.deleted.push(child.id)

		const result = reconcile(SAMPLE_IR, reading)
		const after = result.ir.slides[0]?.nodes.find((node) => node.id === group.id)
		if (after?.kind !== 'group') throw new Error('the group was removed with its child')
		expect(after.children.map((entry) => entry.id)).not.toContain(child.id)
		expect(result.slides[0]?.edits).toBe(1)
	})
})
