import { describe, expect, it } from 'vitest'
import { CORPUS, corpusBytes } from '../corpus/decks'
import { assertNonTrivial, viewDeck } from './roundtrip'

// The corpus is only useful if it is real. Two ways it could be fake and still
// green: a deck that failed to build and yielded nothing, and a deck that built
// but whose content the read model cannot see (so every later comparison is
// empty-against-empty). Both are checked here, once, before any lane runs.

describe('corpus', () => {
	it('names are unique — the snapshot and the coverage report key on them', () => {
		expect(new Set(CORPUS.map((entry) => entry.name)).size).toBe(CORPUS.length)
	})

	for (const entry of CORPUS) {
		it(`${entry.name} builds into a readable, non-empty deck`, async () => {
			assertNonTrivial(await viewDeck(await corpusBytes(entry)), entry.name)
		})
	}
})
