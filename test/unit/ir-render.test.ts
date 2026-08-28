import { describe, expect, it } from 'vitest'
import {
	cellNodeId,
	domNodeId,
	EMU_PER_INCH,
	emuOf,
	IR_VERSION,
	importedNodeId,
	inchesOf,
	NODE_KINDS,
	type NodeId,
	type RenderNode,
} from '../../src/ir/render'
import { SAMPLE_IR } from '../fixtures/render-ir'

/** Every node in the deck, groups flattened, in document order. */
function allNodes(nodes: readonly RenderNode[]): RenderNode[] {
	return nodes.flatMap((node) => (node.kind === 'group' ? [node, ...allNodes(node.children)] : [node]))
}

const NODES = SAMPLE_IR.slides.flatMap((slide) => allNodes(slide.nodes))

describe('the IR survives the JSON island', () => {
	it('is identical after a JSON round trip', () => {
		// The renderer embeds this model in the document and the parser reads it
		// back. `toStrictEqual` rather than `toEqual` on purpose: it distinguishes
		// a missing key from a key holding `undefined`, which is exactly the
		// distinction JSON erases and the one the model forbids.
		// The JSON round trip is the assertion, not a way to copy: `structuredClone`
		// here would prove that the model survives structured cloning, which is not
		// the claim this test makes.
		// oxlint-disable-next-line unicorn/prefer-structured-clone
		expect(JSON.parse(JSON.stringify(SAMPLE_IR))).toStrictEqual(SAMPLE_IR)
	})

	it('holds only JSON-native values', () => {
		// The round-trip test above passes for a `Uint8Array` too — it just comes
		// back as `{"0":…}` and compares unequal only if something reads it. This
		// walks the model instead, so a `Date`, a `Map`, or inline media bytes are
		// caught at the point they are introduced rather than on the way back.
		const offenders: string[] = []
		walk(SAMPLE_IR, '$', offenders)
		expect(offenders).toEqual([])
	})
})

function walk(value: unknown, path: string, offenders: string[]): void {
	if (value === null || typeof value === 'boolean' || typeof value === 'string') return
	if (typeof value === 'number') {
		if (!Number.isFinite(value)) offenders.push(`${path}: ${value}`)
		return
	}
	if (Array.isArray(value)) {
		for (const [index, entry] of value.entries()) walk(entry, `${path}[${index}]`, offenders)
		return
	}
	if (typeof value !== 'object') {
		offenders.push(`${path}: ${typeof value}`)
		return
	}
	if (Object.getPrototypeOf(value) !== Object.prototype) {
		offenders.push(`${path}: ${value.constructor?.name ?? 'exotic object'}`)
		return
	}
	for (const [key, entry] of Object.entries(value)) walk(entry, `${path}.${key}`, offenders)
}

describe('node identity', () => {
	it('is derived from the source, so a second import agrees with the first', () => {
		// Structural, not hashed: the pair is already unique within a deck, already
		// stable, and reads back in a stack trace. A hash would buy opacity and
		// cost collisions.
		expect(importedNodeId(3, 7)).toBe('s3.sp7')
		expect(importedNodeId(3, 7)).toBe(importedNodeId(3, 7))
		expect(cellNodeId(importedNodeId(3, 7), 1, 2)).toBe('s3.sp7.r1c2')
		expect(domNodeId(1, [0, 3, 2])).toBe('s1.dom-0-3-2')
		expect(domNodeId(1, [])).toBe('s1.dom')
	})

	it('is unique across the deck, including table cells and group children', () => {
		const ids: NodeId[] = NODES.map((node) => node.id)
		for (const node of NODES) {
			if (node.kind === 'table') for (const row of node.rows) for (const cell of row.cells) ids.push(cell.id)
		}
		expect(new Set(ids).size).toBe(ids.length)
	})
})

describe('the fixture covers the model', () => {
	it('exercises every node kind', () => {
		// The fixture is what the JSON and surface tests run against, so a node
		// kind missing from it is a kind nothing checks. Adding a kind to
		// `NODE_KINDS` therefore fails here until the fixture grows one.
		expect(new Set(NODES.map((node) => node.kind))).toEqual(new Set(NODE_KINDS))
	})

	it('declares the IR version it was written for', () => {
		expect(SAMPLE_IR.irVersion).toBe(IR_VERSION)
	})
})

describe('units', () => {
	it('converts inches to the integer EMU the format stores', () => {
		expect(emuOf(1)).toBe(EMU_PER_INCH)
		expect(emuOf(0.5)).toBe(457_200)
		expect(Number.isInteger(emuOf(1 / 3))).toBe(true)
		expect(inchesOf(emuOf(2.5))).toBe(2.5)
	})

	it('keeps every frame in the deck integral', () => {
		// EMU is the stored truth precisely so that no frame carries float noise;
		// a fractional one means something converted through inches and back.
		for (const node of NODES) {
			if (node.placement === null) continue
			const { x, y, w, h } = node.placement.box
			expect([x, y, w, h].every(Number.isInteger)).toBe(true)
		}
	})
})
