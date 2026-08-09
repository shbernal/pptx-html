/**
 * The preset catalogue and the path conversion.
 *
 * Hand-built geometry rather than corpus decks, because the interesting cases
 * are the ones the corpus cannot reach: a preset with no local formula, an
 * `arcTo` (which the writer's freeform DSL emits but none of the generated decks
 * use), and a path whose unit space is not the shape's.
 */

import { describe, expect, it } from 'vitest'
import type { Geometry } from '../../src/ir/render'
import { pathOf } from '../../src/render/geometry'

describe('presets', () => {
	it('draws a bounding box for a preset it has no formula for, and says so', () => {
		// The honest half of an incomplete catalogue. A wrong outline drawn
		// confidently is worse than a box, because only one of the two looks like an
		// error — so the fallback is reported and the caller marks the DOM with it.
		const geometry: Geometry = { kind: 'preset', preset: 'wedgeRoundRectCallout', adjustValues: {} }
		const result = pathOf(geometry, 1000, 500)
		expect(result.fallback).toBe('wedgeRoundRectCallout')
		expect(result.d).toBe('M 0 0 L 1000 0 L 1000 500 L 0 500 Z')
	})

	it('reports no fallback for one it draws', () => {
		expect(pathOf({ kind: 'preset', preset: 'ellipse', adjustValues: {} }, 100, 50).fallback).toBeUndefined()
	})

	it('honours an adjust value and falls back to the preset default without one', () => {
		// `adj` is hundred-thousandths of the shorter side, so 25000 on a 400-tall
		// box is a 100 radius — and the default (16667) is a different, smaller one.
		const stated = pathOf({ kind: 'preset', preset: 'roundRect', adjustValues: { adj: 'val 25000' } }, 1000, 400)
		expect(stated.d.startsWith('M 100 0')).toBe(true)

		const defaulted = pathOf({ kind: 'preset', preset: 'roundRect', adjustValues: {} }, 1000, 400)
		expect(defaulted.d.startsWith('M 66.668 0')).toBe(true)
	})

	it('ignores a computed guide rather than parsing half of it', () => {
		// `a:avLst` guides are a formula language. Reading the literal form and
		// silently mis-reading everything else would turn an unmodeled adjust handle
		// into a wrong shape drawn with confidence; falling back to the default is
		// the same thing that happens when the attribute is absent.
		const computed = pathOf({ kind: 'preset', preset: 'roundRect', adjustValues: { adj: '*/ 100 w ss' } }, 1000, 400)
		const absent = pathOf({ kind: 'preset', preset: 'roundRect', adjustValues: {} }, 1000, 400)
		expect(computed.d).toBe(absent.d)
	})

	it('clamps a radius that would exceed the box', () => {
		// `val 100000` is the whole shorter side, and a radius of that would make the
		// two corner arcs on one edge overlap and the path self-intersect.
		const result = pathOf({ kind: 'preset', preset: 'roundRect', adjustValues: { adj: 'val 100000' } }, 1000, 400)
		expect(result.d.startsWith('M 200 0')).toBe(true)
	})
})

describe('custom geometry', () => {
	it('scales path units into the shape box', () => {
		// The path states its own denominator, so the same commands in a 0..100 space
		// and a 0..1000 space describe the same outline at different precisions.
		const geometry: Geometry = {
			kind: 'custom',
			paths: [
				{
					w: 100,
					h: 100,
					fill: 'norm',
					stroke: true,
					commands: [{ cmd: 'moveTo', x: 0, y: 0 }, { cmd: 'lnTo', x: 100, y: 50 }, { cmd: 'close' }],
				},
			],
		}
		expect(pathOf(geometry, 2000, 400).d).toBe('M 0 0 L 2000 200 Z')
	})

	it('treats a zero denominator as "already in shape space"', () => {
		// `@w`/`@h` default to 0, and there is no other sane reading of a divisor of
		// zero — scaling by it would put every command at infinity.
		const geometry: Geometry = {
			kind: 'custom',
			paths: [{ w: 0, h: 0, fill: 'norm', stroke: true, commands: [{ cmd: 'moveTo', x: 10, y: 20 }] }],
		}
		expect(pathOf(geometry, 500, 500).d).toBe('M 10 20')
	})

	it('converts an arcTo from start/sweep angles to an SVG end point', () => {
		// The one command whose conversion is not a rename: OOXML states an arc by
		// where it starts and how far it sweeps, SVG by where it ends. A quarter turn
		// clockwise from the pen at (100, 0), on radii of 100, ends at (200, 100).
		const geometry: Geometry = {
			kind: 'custom',
			paths: [
				{
					w: 0,
					h: 0,
					fill: 'norm',
					stroke: true,
					commands: [
						{ cmd: 'moveTo', x: 100, y: 0 },
						{ cmd: 'arcTo', wR: 100, hR: 100, stAng: -90, swAng: 90 },
					],
				},
			],
		}
		const [, arc] = pathOf(geometry, 1, 1).d.split(' M ').join('M ').split('M 100 0 ')
		expect(arc).toMatch(/^A 100 100 0 0 1 200 100$/)
	})

	it('does not fill a freeform whose every path says it is unfilled', () => {
		// `@fill="none"` marks an open stroke. Filling it would paint an opaque
		// region over whatever the line was drawn across.
		const geometry: Geometry = {
			kind: 'custom',
			paths: [
				{
					w: 0,
					h: 0,
					fill: 'none',
					stroke: true,
					commands: [
						{ cmd: 'moveTo', x: 0, y: 0 },
						{ cmd: 'lnTo', x: 10, y: 10 },
					],
				},
			],
		}
		expect(pathOf(geometry, 100, 100).unfilled).toBe(true)
	})

	it('falls back to a box when the geometry has no drawable command', () => {
		const geometry: Geometry = { kind: 'custom', paths: [{ w: 10, h: 10, fill: 'norm', stroke: true, commands: [] }] }
		expect(pathOf(geometry, 100, 50)).toStrictEqual({ d: 'M 0 0 L 100 0 L 100 50 L 0 50 Z', fallback: 'custGeom' })
	})
})
