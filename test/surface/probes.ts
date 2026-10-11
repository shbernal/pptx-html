/**
 * The probe manifest: writer options sent through the loop one at a time.
 *
 * Each entry is one option set to one value on an otherwise plain deck (see
 * `probe.ts`). The coverage loop grows this file one probe per commit; the
 * frontier it works from is `pnpm run oracle:score`, which lists every inventory
 * path that neither a probe nor `SKIPPED` accounts for.
 *
 * Rules a probe must keep, each of which a test enforces or the score exposes:
 *
 * - **Change one thing.** `author(pptx, false)` and `author(pptx, true)` differ by
 *   the option under test and nothing else, or the signature measures the wrong
 *   thing. Prefer the helpers below, which hold that by construction.
 * - **Use a non-default value.** A default writes nothing and the probe is inert.
 * - **Never edit a probe to make it pass.** A loss is recorded with `issue` (the
 *   upstream lane loses it) or `accepted` (with a reason), never by changing the
 *   value until it happens to survive.
 */

import type TsPptx from 'pptx-ts'
import type { ShapeProps, TableCellProps, TableProps, TextPropsOptions } from 'pptx-ts'
import type { Probe } from './probe'

const BOX = { x: 1, y: 1, w: 5, h: 1.5 }

/** `addText` on one box. `options` is applied only when the probe is on. */
function text(id: string, options: TextPropsOptions, covers = [id.split(':')[0] as string]): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addText('Probe text', { ...BOX, objectName: 'probe', ...(on ? options : {}) })
		},
	}
}

/** `addShape` with a visible fill, so geometry and line options have something to sit on. */
function shape(id: string, options: ShapeProps, covers = [id.split(':')[0] as string]): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			pptx
				.addSlide()
				.addShape('rect', { ...BOX, fill: { color: 'DDE3F0' }, objectName: 'probe', ...(on ? options : {}) })
		},
	}
}

/** A 2×2 table; `options` go on the table call. */
function table(id: string, options: TableProps, covers = [id.split(':')[0] as string]): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addTable(
				[
					[{ text: 'a' }, { text: 'b' }],
					[{ text: 'c' }, { text: 'd' }],
				],
				{ ...BOX, objectName: 'probe', ...(on ? options : {}) }
			)
		},
	}
}

/** A 2×2 table; `options` go on the first cell only. */
function cell(id: string, options: TableCellProps, covers = [id.split(':')[0] as string]): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addTable(
				[
					[{ text: 'a', options: on ? options : {} }, { text: 'b' }],
					[{ text: 'c' }, { text: 'd' }],
				],
				{ ...BOX, objectName: 'probe' }
			)
		},
	}
}

export const PROBES: Probe[] = [
	text('addText.bold', { bold: true }),
	text('addText.italic', { italic: true }),
	text('addText.underline.style:dbl', { underline: { style: 'dbl' } }),
	text('addText.strike:dblStrike', { strike: 'dblStrike' }),
	{ ...text('addText.charSpacing', { charSpacing: 4 }), issue: 'https://github.com/shbernal/pptx-ts/issues/53' },
	text('addText.lineSpacingMultiple', { lineSpacingMultiple: 1.5 }),
	text('addText.valign:bottom', { valign: 'bottom' }),
	text('addText.margin', { margin: [10, 20, 10, 20] }),
	shape('addShape.line.dashType:dash', { line: { color: '250F6B', width: 2, dashType: 'dash' } }, [
		'addShape.line.dashType',
	]),
	shape('addShape.rotate', { rotate: 30 }),
	table('addTable.border.color', { border: { type: 'solid', color: '9B1B30', width: 2 } }, ['addTable.border.color']),
	cell('addTable.cell.fill.color', { fill: { color: 'F0C808' } }),
]

/**
 * Inventory paths deliberately not probed, and why. A path here counts as covered,
 * so the reason has to hold up: "hard to probe" is not one; "selects the input
 * source and writes nothing of its own" is.
 */
export const SKIPPED: Record<string, string> = {
	'addText.data': 'an image source for picture-filled text, exercised by addText.fill.image.*',
	'addText.path': 'an image source for picture-filled text, exercised by addText.fill.image.*',
}
