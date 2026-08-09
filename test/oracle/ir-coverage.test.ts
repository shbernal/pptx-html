import { describe, expect, it } from 'vitest'
import { importDeck } from '../../src/import/deck'
import { NODE_KINDS, type NodeKind, type RenderNode } from '../../src/ir/render'
import { CORPUS, corpusBytes } from '../corpus/decks'

/**
 * Which corpus decks exercise each node kind.
 *
 * Before the importer existed this could only assert that a kind *named* a deck,
 * because nothing built a `RenderIr` from a package. Now it does, so every claim
 * below is checked against what that deck actually produces — a kind whose decks stopped
 * producing it fails here rather than sitting in the model unexercised, and a
 * claim that was always wrong cannot survive being written down.
 */
const EXERCISED_BY: Record<NodeKind, string[]> = {
	shape: [
		'text-box',
		'autoshape',
		'custgeom',
		'gradient',
		'line-cap',
		'bullet',
		'layout-placeholder',
		'rotated-flipped',
		'theme-color',
	],
	picture: ['picture'],
	connector: ['connector'],
	table: ['table'],
	group: ['group', 'nested-group'],
	// No writer emits a chart this model can draw, which is the whole point of the
	// kind: the chart round-trips through `addChart` and renders as a labelled box.
	opaque: ['chart'],
}

/** Every node in a deck, groups flattened. */
function flatten(nodes: readonly RenderNode[]): RenderNode[] {
	return nodes.flatMap((node) => (node.kind === 'group' ? [node, ...flatten(node.children)] : [node]))
}

const kindsByDeck = new Map<string, Set<NodeKind>>()

async function kindsOf(name: string): Promise<Set<NodeKind>> {
	const hit = kindsByDeck.get(name)
	if (hit) return hit
	const entry = CORPUS.find((candidate) => candidate.name === name)
	if (!entry) throw new Error(`the corpus has no deck named ${name}`)
	const { render } = await importDeck(await corpusBytes(entry))
	const kinds = new Set(render.slides.flatMap((slide) => flatten(slide.nodes)).map((node) => node.kind))
	kindsByDeck.set(name, kinds)
	return kinds
}

describe('every node kind has a deck behind it', () => {
	it('claims a corpus deck for each kind', () => {
		expect(Object.keys(EXERCISED_BY).sort()).toEqual([...NODE_KINDS].sort())
		for (const kind of NODE_KINDS) expect(EXERCISED_BY[kind].length).toBeGreaterThan(0)
	})

	for (const kind of NODE_KINDS) {
		for (const name of EXERCISED_BY[kind]) {
			it(`${name} produces a ${kind} node`, async () => {
				expect([...(await kindsOf(name))]).toContain(kind)
			})
		}
	}

	it('no corpus deck produces a kind nothing claims', async () => {
		// The other direction: a deck quietly gaining a node kind means the model
		// grew a producer nobody recorded, and the table above stops describing the
		// corpus.
		const claimed = new Set(Object.keys(EXERCISED_BY))
		const produced = new Set<string>()
		for (const entry of CORPUS) for (const kind of await kindsOf(entry.name)) produced.add(kind)
		expect([...produced].filter((kind) => !claimed.has(kind))).toEqual([])
	})
})
