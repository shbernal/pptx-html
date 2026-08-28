/**
 * Table cell geometry, on a fixture built for one question each.
 *
 * `renderTable` is not exported, so these go through `renderDeck` and read the
 * cell rectangles back out of the SVG. That is the right level anyway: the thing
 * under test is what a viewer sees, and a cell drawn at the wrong height is
 * invisible to every other layer. This is the visual channel, so nothing here
 * touches the island or the round trip.
 */

import { describe, expect, it } from 'vitest'
import { EMU_PER_INCH, type RenderIr, type TableNode, cellNodeId, importedNodeId } from '../../src/ir/render'
import { renderDeck } from '../../src/render/document'

const NO_ASSETS = { assets: 'ref' } as const

const inch = (n: number): number => Math.round(n * EMU_PER_INCH)

const TABLE_ID = importedNodeId(1, 1)
const NO_BORDERS = {
	left: { kind: 'inherit' },
	right: { kind: 'inherit' },
	top: { kind: 'inherit' },
	bottom: { kind: 'inherit' },
} as const
const NO_MARGINS = { left: null, right: null, top: null, bottom: null }

/** One cell, with only the fields the geometry reads varying. */
function cell(row: number, column: number, rest: { span?: { columns: number; rows: number }; covered?: boolean } = {}) {
	return {
		id: cellNodeId(TABLE_ID, row, column),
		text: null,
		fill: { kind: 'solid', color: { kind: 'srgb', hex: '451DC7' } } as const,
		borders: NO_BORDERS,
		span: rest.span ?? null,
		covered: rest.covered ?? false,
		marginsEmu: NO_MARGINS,
		anchor: 'top' as const,
	}
}

/**
 * A two-row table whose rows are deliberately *unequal*: a header plus a body
 * row, which is the shape most real tables have and the one a span bug hides in.
 */
function unequalRows(span: { columns: number; rows: number }): TableNode {
	return {
		kind: 'table',
		id: TABLE_ID,
		name: 'Table 1',
		placement: {
			box: { x: 0, y: 0, w: inch(4), h: inch(3) },
			rotation: 0,
			flipH: false,
			flipV: false,
			geometrySource: 'own',
		},
		render: 'drawn',
		columns: [{ widthEmu: inch(2) }, { widthEmu: inch(2) }],
		rows: [
			{ heightEmu: inch(0.5), cells: [cell(0, 0, { span }), cell(0, 1)] },
			{ heightEmu: inch(2.5), cells: [cell(1, 0, { covered: true }), cell(1, 1)] },
		],
	}
}

function irWith(table: TableNode): RenderIr {
	return {
		irVersion: 7,
		size: { w: inch(10), h: inch(5.625) },
		assets: [],
		slides: [
			{
				number: 1,
				source: 'authored',
				layout: null,
				hidden: false,
				background: { source: 'master', fill: { kind: 'solid', color: { kind: 'srgb', hex: 'FFFFFF' } } },
				chrome: [],
				nodes: [table],
				notes: null,
				residual: null,
				fidelity: [],
			},
		],
	}
}

/**
 * The `height` of one cell's rectangle. Found through the cell's own
 * `data-pxh-node`, because the slide surface is a `<rect>` too and matching the
 * first one in the document would measure the background.
 */
function cellHeight(html: string, row: number, column: number): number {
	const id = cellNodeId(TABLE_ID, row, column)
	const match = new RegExp(`data-pxh-node="${id}"[^>]*><rect[^>]*height="(\\d+(?:\\.\\d+)?)"`).exec(html)
	if (match?.[1] === undefined) throw new Error(`no rectangle for cell ${id}`)
	return Number(match[1])
}

describe('a cell that spans rows is as tall as the rows it covers', () => {
	it('sums the spanned rows rather than multiplying the first', async () => {
		// 0.5in + 2.5in = 3in. Multiplying the first row's height by the span count
		// would give 1in, which is the bug this pins: it is only the same answer when
		// every row is the same height, and a header row is exactly when it is not.
		const { html } = await renderDeck(irWith(unequalRows({ columns: 1, rows: 2 })), NO_ASSETS)
		expect(cellHeight(html, 0, 0)).toBe(inch(3))
	})

	it('leaves an unspanned cell at its own row height', async () => {
		const { html } = await renderDeck(irWith(unequalRows({ columns: 2, rows: 1 })), NO_ASSETS)
		expect(cellHeight(html, 0, 0)).toBe(inch(0.5))
	})
})
