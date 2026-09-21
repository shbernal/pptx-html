/**
 * Where a node is painted, as opposed to what is painted in it.
 *
 * One rule, stated by the model rather than by the renderer: a child's
 * {@link Placement} is already slide-absolute, so a group is walked for paint
 * order and identity and never composes a transform ({@link GroupNode}). The
 * renderer put the group's own `translate` on the wrapper anyway, so every
 * grouped shape was offset twice and the further a group sat from the origin the
 * further its contents left it.
 *
 * Both directions are here, because the tempting fix is structural -- "a `<g>`
 * whose children are node-addressed does not place them" -- and that is wrong: a
 * table nests node-addressed `<g>`s too, and its cells ARE placed against it.
 */

import { describe, expect, it } from 'vitest'
import type { GroupNode, Placement, RenderNode, ShapeNode, Stroke, TableNode } from '../../src/ir/render'
import { cellNodeId } from '../../src/ir/render'
import { renderNode, type NodeContext } from '../../src/render/node'
import { Defs } from '../../src/render/paint'

const placement = (x: number, y: number, w = 1000, h = 500): Placement => ({
	box: { x, y, w, h },
	rotation: 0,
	flipH: false,
	flipV: false,
	geometrySource: 'own',
})

const shape = (id: string, x: number, y: number): ShapeNode => ({
	kind: 'shape',
	id,
	name: id,
	render: 'drawn',
	placement: placement(x, y),
	geometry: { kind: 'preset', preset: 'rect', adjustValues: {} },
	fill: { kind: 'none' },
	stroke: { kind: 'none' },
	text: null,
})

const groupOf = (id: string, x: number, y: number, children: RenderNode[]): GroupNode => ({
	kind: 'group',
	id,
	name: id,
	render: 'drawn',
	placement: placement(x, y),
	children,
})

const context = (): NodeContext => ({ defs: new Defs(), skipped: [], editable: true })

/** No stroke, for the six edges a plain cell states nothing on. */
const none: Stroke = { kind: 'none' }

/** The `transform` attribute of each `<g>`, in document order; `null` where there is none. */
function transforms(svg: string): Array<string | null> {
	return [...svg.matchAll(/<g\b([^>]*)>/g)].map((match) => /transform="([^"]*)"/.exec(match[1] ?? '')?.[1] ?? null)
}

describe('a group', () => {
	it('places none of its children, because they are already placed', () => {
		const group = groupOf('g1', 4000, 2000, [shape('g1.a', 4000, 2000), shape('g1.b', 5000, 2500)])

		const [outer, first, second] = transforms(renderNode(group, context()))

		// The group draws nothing and offsets nothing; its wrapper stays only to carry the id.
		expect(outer).toBeNull()
		// Each child keeps its own slide-absolute placement, unchanged by the group above it.
		expect(first).toBe('translate(4000 2000)')
		expect(second).toBe('translate(5000 2500)')
	})

	it('still carries its identity, so an edit can address it', () => {
		const group = groupOf('g1', 4000, 2000, [shape('g1.a', 4000, 2000)])

		expect(renderNode(group, context())).toContain('data-pxh-node="g1"')
	})

	it('nests without compounding, so a child of a child is where the model put it', () => {
		const inner = groupOf('g1.inner', 6000, 3000, [shape('g1.inner.a', 6000, 3000)])
		const outer = groupOf('g1', 4000, 2000, [inner])

		// Two groups deep is where the old behaviour was worst: the leaf took both
		// offsets on top of its own.
		expect(transforms(renderNode(outer, context()))).toEqual([null, null, 'translate(6000 3000)'])
	})
})

describe('a table', () => {
	it('places its cells against itself, which is the case the group rule must not catch', () => {
		const table: TableNode = {
			kind: 'table',
			id: 't1',
			name: 't1',
			render: 'drawn',
			placement: placement(2000, 1000, 4000, 2000),
			columns: [{ widthEmu: 4000 }],
			rows: [
				{
					heightEmu: 2000,
					cells: [
						{
							id: cellNodeId('t1', 0, 0),
							text: null,
							fill: { kind: 'none' },
							borders: { left: none, right: none, top: none, bottom: none, tlToBr: none, blToTr: none },
							span: null,
							covered: false,
							marginsEmu: { left: null, right: null, top: null, bottom: null },
							anchor: 'top',
						},
					],
				},
			],
		}

		// The table's own `<g>` keeps its transform: unlike a group, it is what its
		// cells are positioned against.
		expect(transforms(renderNode(table, context()))[0]).toBe('translate(2000 1000)')
	})
})
