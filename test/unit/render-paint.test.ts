/**
 * Preset pattern fills, drawn from their tiles.
 *
 * Read off the `<pattern>` the paint registers rather than off a rendered slide:
 * the question is which pixels of the eight-by-eight tile carry the foreground,
 * and the def is where that is stated once.
 */

import { describe, expect, it } from 'vitest'
import type { Fill } from '../../src/ir/render'
import { Defs, fillPaint } from '../../src/render/paint'

const FRAME = { w: 1000, h: 1000, rotation: 0 }
const RED = { kind: 'srgb', hex: 'C00000' } as const
const CREAM = { kind: 'srgb', hex: 'FFF2CC' } as const

function paint(fill: Fill): { attrs: string; approx?: string; defs: string } {
	const defs = new Defs()
	const painted = fillPaint(fill, FRAME, defs)
	return {
		attrs: painted.attrs,
		...(painted.approx === undefined ? {} : { approx: painted.approx }),
		defs: defs.render(),
	}
}

describe('a preset pattern', () => {
	it('paints its tile: the background under the foreground pixels', () => {
		// `horz` is one full row at the top of the tile and nothing else.
		const { attrs, approx, defs } = paint({ kind: 'pattern', preset: 'horz', foreground: RED, background: CREAM })
		expect(attrs).toBe('fill="url(#pxh-p0)"')
		expect(approx).toBeUndefined()
		expect(defs).toContain('<rect width="76200" height="76200" fill="#FFF2CC"/>')
		expect(defs).toContain('d="M0 0h8v1h-8z" fill="#C00000"')
	})

	it('reads the tile left to right from the high bit, top row first', () => {
		// `dnDiag` falls to the right, so its first pixel is the top-left one.
		const { defs } = paint({ kind: 'pattern', preset: 'dnDiag', foreground: RED, background: CREAM })
		expect(defs).toContain('d="M0 0h1v1h-1zM1 1h1v1h-1zM2 2h1v1h-1z')
	})

	it('falls back to the background colour, marked, for a preset it has no tile for', () => {
		const { attrs, approx } = paint({ kind: 'pattern', preset: 'notAPreset', foreground: RED, background: CREAM })
		expect(attrs).toBe('fill="#FFF2CC"')
		expect(approx).toBe('fill:pattern')
	})
})
