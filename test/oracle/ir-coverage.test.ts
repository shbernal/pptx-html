import { describe, expect, it } from 'vitest'
import { NODE_KINDS, type NodeKind } from '../../src/ir/render'
import { CORPUS } from '../corpus/decks'

/**
 * Which corpus decks will exercise each node kind once the import layer exists.
 *
 * This is a claim, not yet a measurement — nothing builds a `RenderIr` from a
 * package until part 04 lands. What it buys today is that a node kind cannot be
 * added to the model without someone deciding which deck proves it, and a corpus
 * deck cannot be renamed without this failing. When the importer arrives, the
 * assertion tightens from "a deck is named" to "that deck actually produces this
 * kind".
 */
const EXERCISED_BY: Record<NodeKind, string[]> = {
	shape: ['text-box', 'autoshape', 'custgeom', 'gradient', 'layout-placeholder', 'rotated-flipped', 'theme-color'],
	picture: ['picture'],
	connector: ['connector'],
	table: ['table'],
	group: ['group'],
	// No writer emits a chart the render model can draw, which is the whole point
	// of the kind: the slide is carried and the chart renders as a labelled box.
	opaque: ['chart'],
}

describe('every node kind has a deck behind it', () => {
	it('claims a corpus deck for each kind', () => {
		expect(Object.keys(EXERCISED_BY).sort()).toEqual([...NODE_KINDS].sort())
		for (const kind of NODE_KINDS) expect(EXERCISED_BY[kind].length).toBeGreaterThan(0)
	})

	it('names only decks the corpus actually has', () => {
		const known = new Set(CORPUS.map((entry) => entry.name))
		const missing = Object.values(EXERCISED_BY)
			.flat()
			.filter((name) => !known.has(name))
		expect(missing).toEqual([])
	})
})
