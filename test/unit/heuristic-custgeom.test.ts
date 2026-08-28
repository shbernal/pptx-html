/**
 * custGeom emit round-trip, headless. Feeds hand-written `path` IR items through
 * `heuristic/custgeom`, serializes on `@shbernal/ts-pptx`, then parses the deck back
 * with ts-pptx's `read` model and asserts the freeform geometry survives. This is
 * where custGeom fidelity is proven without a browser.
 */

import { ShapeType, TsPptx } from '@shbernal/ts-pptx'
import { type AnyShape, type AutoShape, type CustomGeometry, isAutoShape, Presentation } from '@shbernal/ts-pptx/read'
import { describe, expect, it } from 'vitest'
import { pathShapeOptions } from '../../src/heuristic/custgeom'
import type { PathItem, SlideModel } from '../../src/heuristic/model'
import { addModelToSlide } from '../../src/heuristic/slide'
import { first, only } from '../support'

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

// `toBytes()` is what `Presentation.load` takes, so the deck never becomes base64
// on the way from the writer to the reader.
async function emit(items: PathItem[]) {
	const pptx = new TsPptx()
	pptx.layout = 'LAYOUT_16x9'
	const slide = pptx.addSlide()
	const issues = await addModelToSlide({ ShapeType }, slide, model(items), SIZE)
	const bytes = await pptx.toBytes()
	return { issues, bytes }
}

type FreeformShape = AutoShape & { customGeometry: CustomGeometry }

/**
 * The freeform shapes on a slide. `AnyShape` also covers connectors and pictures,
 * which carry no geometry at all, so the guard is what makes `customGeometry`
 * readable — and a run where the writer stopped emitting autoshapes fails here
 * rather than reading `undefined` off the wrong shape kind.
 */
function freeforms(shapes: AnyShape[]): FreeformShape[] {
	return shapes.filter((shape): shape is FreeformShape => isAutoShape(shape) && shape.customGeometry !== null)
}

/** The single freeform on a slide, asserted present so a miss fails loudly. */
function onlyFreeform(shapes: AnyShape[]): FreeformShape {
	return only(freeforms(shapes), 'freeform shape')
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

describe('emit → ts-pptx → read round-trip (custGeom)', () => {
	it('emits path items without per-item issues', async () => {
		const { issues } = await emit([triangle(), cubicStroke()])
		expect(issues).toEqual([])
	})

	it('writes custom geometry (not preset) shapes ts-pptx can read back', async () => {
		const { bytes } = await emit([triangle(), cubicStroke()])
		const pres = await Presentation.load(bytes)
		const customs = freeforms(first(pres.slides, 'slide').shapes)
		expect(customs.length).toBe(2)
		for (const shape of customs) {
			expect(shape.presetGeometry).toBeNull()
		}
	})

	it('round-trips the triangle command sequence (moveTo / lnTo×2 / close)', async () => {
		const { bytes } = await emit([triangle()])
		const pres = await Presentation.load(bytes)
		const shape = onlyFreeform(first(pres.slides, 'slide').shapes)
		const cmds = first(shape.customGeometry.paths, 'path').commands.map((command) => command.cmd)
		expect(cmds).toEqual(['moveTo', 'lnTo', 'lnTo', 'close'])
	})

	it('preserves the cubic curve as a cubicBezTo segment', async () => {
		const { bytes } = await emit([cubicStroke()])
		const pres = await Presentation.load(bytes)
		const shape = onlyFreeform(first(pres.slides, 'slide').shapes)
		const cmds = first(shape.customGeometry.paths, 'path').commands.map((command) => command.cmd)
		expect(cmds).toEqual(['moveTo', 'cubicBezTo'])
	})

	it('scales path-unit coordinates to the shape box (triangle apex at half width, full height)', async () => {
		const { bytes } = await emit([triangle()])
		const pres = await Presentation.load(bytes)
		const shape = onlyFreeform(first(pres.slides, 'slide').shapes)
		const path = first(shape.customGeometry.paths, 'path')
		// Apex point (box-inch 1,2 in a 2×2 box) → path-units (w/2, h).
		// The command list is a union discriminated on `cmd`; only the line and move
		// verbs carry a coordinate, so the narrowing is what makes `x`/`y` readable.
		const apex = path.commands[2] // third command = second lnTo
		if (apex?.cmd !== 'lnTo') throw new Error(`expected an lnTo at index 2, got ${apex?.cmd ?? 'nothing'}`)
		expect(apex.x).toBeCloseTo(path.w / 2, -2)
		expect(apex.y).toBeCloseTo(path.h, -2)
	})
})
