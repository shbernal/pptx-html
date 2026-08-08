import { describe, expect, it } from 'vitest'
import { CORPUS, corpusBytes } from '../corpus/decks'
import { assertNonTrivial, viewDeck, weigh } from './roundtrip'

// The corpus is only useful if it is real. Two ways it could be fake and still
// green: a deck that failed to build and yielded nothing, and a deck that built
// but whose content the read model cannot see (so every later comparison is
// empty-against-empty). Both are checked here, once, before any lane runs.

describe('corpus', () => {
	it('has both tiers', () => {
		expect(CORPUS.filter((entry) => entry.tier === 'primitive').length).toBeGreaterThan(0)
		expect(CORPUS.filter((entry) => entry.tier === 'hard').length).toBeGreaterThan(0)
	})

	it('names are unique — the snapshot and the coverage report key on them', () => {
		expect(new Set(CORPUS.map((entry) => entry.name)).size).toBe(CORPUS.length)
	})

	for (const entry of CORPUS) {
		it(`${entry.name} builds into a readable, non-empty deck`, async () => {
			const bytes = await corpusBytes(entry)
			expect(bytes.byteLength).toBeGreaterThan(0)

			const view = await viewDeck(bytes)
			assertNonTrivial(view, entry.name)
			expect(weigh(view).slides).toBeGreaterThan(0)
		})
	}

	it('generation is deterministic under the normalized model', async () => {
		// Same builder, two runs. Byte equality is explicitly not required — rIds,
		// cNvPr ids and zip metadata may differ — but the canonical model must not.
		const entry = CORPUS[0]
		const first = await viewDeck(await entry.build())
		const second = await viewDeck(await entry.build())
		expect(second.canonical).toEqual(first.canonical)
	})
})
