import { describe, expect, it } from 'vitest'
import {
	cellNodeId,
	EMU_PER_INCH,
	emuOf,
	importedNodeId,
	inchesOf,
	NODE_KINDS,
	type RenderNode,
} from '../../src/ir/render'
import { SAMPLE_IR } from '../fixtures/render-ir'

/** Every node in the deck, groups flattened, in document order. */
function allNodes(nodes: readonly RenderNode[]): RenderNode[] {
	return nodes.flatMap((node) => (node.kind === 'group' ? [node, ...allNodes(node.children)] : [node]))
}

const NODES = SAMPLE_IR.slides.flatMap((slide) => allNodes(slide.nodes))

describe('node identity', () => {
	it('is derived from the source, so a second import agrees with the first', () => {
		// Structural, not hashed: the pair is already unique within a deck, already
		// stable, and reads back in a stack trace. A hash would buy opacity and
		// cost collisions.
		expect(importedNodeId(3, 7)).toBe('s3.sp7')
		expect(cellNodeId(importedNodeId(3, 7), 1, 2)).toBe('s3.sp7.r1c2')
	})
})

describe('the fixture covers the model', () => {
	it('exercises every node kind', () => {
		// The fixture is what the island and surface tests run against, so a node
		// kind missing from it is a kind nothing checks. Adding a kind to
		// `NODE_KINDS` therefore fails here until the fixture grows one.
		expect(new Set(NODES.map((node) => node.kind))).toEqual(new Set(NODE_KINDS))
	})
})

describe('units', () => {
	it('converts inches to the integer EMU the format stores', () => {
		expect(emuOf(1)).toBe(EMU_PER_INCH)
		expect(emuOf(0.5)).toBe(457_200)
		expect(Number.isInteger(emuOf(1 / 3))).toBe(true)
		expect(inchesOf(emuOf(2.5))).toBe(2.5)
	})
})
