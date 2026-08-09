/**
 * The decks the playground offers, taken from the oracle's own corpus.
 *
 * Not a separate set of demo decks, and that is the point: `test/corpus/decks.ts`
 * is the input domain the round-trip gate is scoped to, so a sample shown here is
 * a deck CI actually proves something about. A hand-written demo deck would be
 * free to be one the pipeline happens to be good at.
 *
 * The `hard` tier is shown *as* hard, labelled, because a hard-tier deck presented
 * without its tier would be the site claiming coverage the oracle does not gate.
 */

import { CORPUS, type CorpusDeck } from '../../../../test/corpus/decks.ts'

/**
 * A curated subset, in the order they appear in the picker — the whole corpus
 * would be a wall of near-identical names, and these are the ones whose
 * differences a visitor can see.
 */
export const SAMPLE_NAMES = ['text-box', 'autoshape', 'custgeom', 'table', 'bullet', 'gradient', 'chart'] as const

export interface Sample {
	name: string
	tier: CorpusDeck['tier']
	/** The corpus's own one-line statement of what the deck is there to hold the pipeline to. */
	why: string
	build: () => Promise<Uint8Array>
}

/**
 * Resolve the curated names against the corpus, refusing to drop one it cannot
 * find. A corpus rename must break this rather than quietly shorten the picker —
 * the coupling is the feature, so it has to be a loud one.
 */
export function samples(): Sample[] {
	const missing = SAMPLE_NAMES.filter((name) => !CORPUS.some((deck) => deck.name === name))
	if (missing.length > 0) {
		throw new Error(
			`the playground asks for corpus decks that do not exist: ${missing.join(', ')}. ` +
				'Fix the names in site/.vitepress/theme/playground/samples.ts, or restore them in test/corpus/decks.ts.'
		)
	}
	return SAMPLE_NAMES.map((name) => {
		const deck = CORPUS.find((entry) => entry.name === name)
		if (!deck) throw new Error(`corpus deck ${name} vanished between the check and the lookup`)
		return { name: deck.name, tier: deck.tier, why: deck.why, build: deck.build }
	})
}
