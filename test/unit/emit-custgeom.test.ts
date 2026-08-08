/**
 * custGeom emit round-trip, headless. Feeds hand-written `path` IR items through
 * `emit/custgeom`, serializes on `@shbernal/ts-pptx`, runs the repair pass, then
 * parses the deck back with ts-pptx's `read` model and asserts the freeform
 * geometry survives. This is where custGeom fidelity is proven without a browser.
 */

// @ts-expect-error — ts-pptx ships its own types; node-resolved entry is fine for tests.
import { ShapeType, TsPptx } from '@shbernal/ts-pptx'
// @ts-expect-error — read entry typed via package exports.
import { Presentation } from '@shbernal/ts-pptx/read'
import { describe, expect, it } from 'vitest'
import { pathShapeOptions } from '../../src/emit/custgeom'
import { addModelToSlide } from '../../src/emit/slide'
import type { PathItem, SlideModel } from '../../src/ir/model'
import { repairPptxBase64 } from '../../src/repair/repair'

const SIZE = { width: 13.333, height: 7.5 }

// A filled triangle and a cubic-curve stroke, both authored in box-inch space.
function triangle(): PathItem {
	return {
		type: 'path',
		z: 1,
		position: { x: 1, y: 1, w: 2, h: 2 },
		fill: '451DC7',
		points: [{ x: 0, y: 0, moveTo: true }, { x: 2, y: 0 }, { x: 1, y: 2 }, { close: true }],
	}
}

function cubicStroke(): PathItem {
	return {
		type: 'path',
		z: 2,
		position: { x: 4, y: 1, w: 3, h: 1 },
		line: { color: '04F06A', width: 2 },
		points: [
			{ x: 0, y: 1, moveTo: true },
			{ x: 1.5, y: 0, curve: { type: 'cubic', x1: 0.5, y1: 0, x2: 1, y2: 0 } },
		],
	}
}

function model(items: PathItem[]): SlideModel {
	return { background: { type: 'color', value: 'FFFFFF' }, items }
}

async function emit(items: PathItem[]) {
	const pptx = new TsPptx()
	pptx.layout = 'LAYOUT_16x9'
	const slide = pptx.addSlide()
	const issues = await addModelToSlide({ ShapeType }, slide, model(items), SIZE)
	const base64 = await repairPptxBase64(await pptx.write({ outputType: 'base64' }))
	return { issues, base64 }
}

function loadDeck(base64: string) {
	return Presentation.load(Uint8Array.from(Buffer.from(base64, 'base64')))
}

describe('pathShapeOptions — fill / stroke mapping', () => {
	it('maps a fill color and omits line when there is no stroke', () => {
		const opts = pathShapeOptions(triangle())
		expect(opts.fill).toEqual({ color: '451DC7' })
		expect(opts.line).toBeUndefined()
		expect(opts.x).toBe(1)
		expect(opts.points).toHaveLength(4)
	})

	it('maps a stroke to line and omits fill when there is none', () => {
		const opts = pathShapeOptions(cubicStroke())
		expect(opts.fill).toBeUndefined()
		expect(opts.line).toEqual({ color: '04F06A', width: 2, dashType: 'solid' })
	})
})

describe('emit → ts-pptx → repair → read round-trip (custGeom)', () => {
	it('emits path items without per-item issues', async () => {
		const { issues } = await emit([triangle(), cubicStroke()])
		expect(issues).toEqual([])
	})

	it('writes custom geometry (not preset) shapes ts-pptx can read back', async () => {
		const { base64 } = await emit([triangle(), cubicStroke()])
		const pres = await loadDeck(base64)
		const shapes = pres.slides[0].shapes
		const customs = shapes.filter((s: { customGeometry?: unknown }) => s.customGeometry)
		expect(customs.length).toBe(2)
		for (const s of customs) {
			expect(s.presetGeometry).toBeNull()
		}
	})

	it('round-trips the triangle command sequence (moveTo / lnTo×2 / close)', async () => {
		const { base64 } = await emit([triangle()])
		const pres = await loadDeck(base64)
		const shape = pres.slides[0].shapes.find((s: { customGeometry?: unknown }) => s.customGeometry)
		const cmds = shape.customGeometry.paths[0].commands.map((c: { cmd: string }) => c.cmd)
		expect(cmds).toEqual(['moveTo', 'lnTo', 'lnTo', 'close'])
	})

	it('preserves the cubic curve as a cubicBezTo segment', async () => {
		const { base64 } = await emit([cubicStroke()])
		const pres = await loadDeck(base64)
		const shape = pres.slides[0].shapes.find((s: { customGeometry?: unknown }) => s.customGeometry)
		const cmds = shape.customGeometry.paths[0].commands.map((c: { cmd: string }) => c.cmd)
		expect(cmds).toEqual(['moveTo', 'cubicBezTo'])
	})

	it('scales path-unit coordinates to the shape box (triangle apex at half width, full height)', async () => {
		const { base64 } = await emit([triangle()])
		const pres = await loadDeck(base64)
		const shape = pres.slides[0].shapes.find((s: { customGeometry?: unknown }) => s.customGeometry)
		const path = shape.customGeometry.paths[0]
		// Apex point (box-inch 1,2 in a 2×2 box) → path-units (w/2, h).
		const apex = path.commands[2] // third command = second lnTo
		expect(apex.x).toBeCloseTo(path.w / 2, -2)
		expect(apex.y).toBeCloseTo(path.h, -2)
	})
})
