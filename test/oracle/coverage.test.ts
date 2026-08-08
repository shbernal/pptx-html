import { describe, expect, it } from 'vitest'
import { CORPUS, corpusBytes } from '../corpus/decks'
import { coverageRow, formatCoverage, unmatchableConstructs } from './coverage'
import { viewDeck } from './roundtrip'

// The snapshot below is the support baseline. It is expected to change — that is
// the point — but never silently: a construct sliding from `flattened` to
// `dropped`, or a slide falling into the carried lane, shows up as a diff here
// even when every round-trip stays green.

describe('coverage', () => {
	it('matches the recorded support baseline', async () => {
		const rows = []
		for (const entry of CORPUS) {
			const view = await viewDeck(await corpusBytes(entry))
			rows.push(coverageRow(entry.name, entry.tier, view.ir))
		}
		expect(formatCoverage(rows)).toMatchSnapshot()
	})

	it('every note names a construct the differ can match', async () => {
		// A note the matcher table does not know is inert: it can never declare a
		// difference, so it looks like honesty while excusing nothing.
		const unmatchable = new Set<string>()
		for (const entry of CORPUS) {
			const view = await viewDeck(await corpusBytes(entry))
			for (const construct of unmatchableConstructs(view.ir.fidelity)) unmatchable.add(construct)
		}
		expect([...unmatchable]).toEqual([])
	})
})
