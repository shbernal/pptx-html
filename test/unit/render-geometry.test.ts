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

	it('draws a chevron, taking its inset from the shorter side', () => {
		// `maxAdj` is `100000 * w / ss`, so `adj` is a fraction of the *shortest* side
		// and not of the width. On a wide box that is the whole difference between a
		// chevron and something with a 100-wide notch: 200×100 insets by 50, not 100.
		const square = pathOf({ kind: 'preset', preset: 'chevron', adjustValues: {} }, 100, 100)
		expect(square.d).toBe('M 0 0 L 50 0 L 100 50 L 50 100 L 0 100 L 50 50 Z')
		expect(square.fallback).toBeUndefined()

		const wide = pathOf({ kind: 'preset', preset: 'chevron', adjustValues: {} }, 200, 100)
		expect(wide.d).toBe('M 0 0 L 150 0 L 200 50 L 150 100 L 0 100 L 50 50 Z')
	})

	it('draws the five presets counted in the browser-preview evidence', () => {
		const draw = (preset: string) => pathOf({ kind: 'preset', preset, adjustValues: {} }, 1000, 500)
		for (const preset of ['round2DiagRect', 'rightArrowCallout', 'rightArrow', 'round2SameRect', 'wedgeRectCallout']) {
			expect([preset, draw(preset).fallback]).toStrictEqual([preset, undefined])
		}

		expect(draw('rightArrow').d).toBe('M 0 125 L 750 125 L 750 0 L 1000 250 L 750 500 L 750 375 L 0 375 Z')
		expect(draw('rightArrowCallout').d).toBe(
			'M 0 0 L 649.77 0 L 649.77 187.5 L 875 187.5 L 875 125 L 1000 250 ' +
				'L 875 375 L 875 312.5 L 649.77 312.5 L 649.77 500 L 0 500 Z'
		)
	})

	it('honours both radius handles on the two-corner rectangle presets', () => {
		const diag = pathOf(
			{ kind: 'preset', preset: 'round2DiagRect', adjustValues: { adj1: 'val 20000', adj2: 'val 10000' } },
			1000,
			500
		)
		expect(diag.d).toBe(
			'M 100 0 L 950 0 A 50 50 0 0 1 1000 50 L 1000 400 A 100 100 0 0 1 900 500 ' +
				'L 50 500 A 50 50 0 0 1 0 450 L 0 100 A 100 100 0 0 1 100 0 Z'
		)

		const same = pathOf(
			{ kind: 'preset', preset: 'round2SameRect', adjustValues: { adj1: 'val 20000', adj2: 'val 10000' } },
			1000,
			500
		)
		expect(same.d).toBe(
			'M 100 0 L 900 0 A 100 100 0 0 1 1000 100 L 1000 450 A 50 50 0 0 1 950 500 ' +
				'L 50 500 A 50 50 0 0 1 0 450 L 0 100 A 100 100 0 0 1 100 0 Z'
		)
	})

	it('collapses a wedge callout handle inside the rectangle onto its edges', () => {
		// DR-18-0013 corrects the electronic addendum here. Its original guides put
		// an inward notch at the handle; PowerPoint draws the full rectangle.
		const result = pathOf(
			{ kind: 'preset', preset: 'wedgeRectCallout', adjustValues: { adj1: 'val 0', adj2: 'val 0' } },
			120,
			60
		)
		expect(result.d).toBe(
			'M 0 0 L 20 0 L 20 0 L 50 0 L 120 0 L 120 10 L 120 10 L 120 25 L 120 60 ' +
				'L 50 60 L 20 60 L 20 60 L 0 60 L 0 25 L 0 30 L 0 10 Z'
		)
	})

	it('draws a five-pointed star whose points reach every edge of the box', () => {
		// `hf`/`vf` (105146 and 110557) exist to stretch the pentagon's circumscribed
		// circle until the outer points sit exactly on the box. Asserted as a property
		// rather than only as a string, because dropping either factor still produces
		// a plausible star — just an inset one, which is the failure that would read
		// as correct in a screenshot.
		const result = pathOf({ kind: 'preset', preset: 'star5', adjustValues: {} }, 100, 100)
		const numbers = result.d.match(/-?\d+(\.\d+)?/g)?.map(Number) ?? []
		const xs = numbers.filter((_, index) => index % 2 === 0)
		const ys = numbers.filter((_, index) => index % 2 === 1)
		expect([Math.min(...xs), Math.max(...xs)]).toStrictEqual([0, 100])
		expect([Math.min(...ys), Math.max(...ys)]).toStrictEqual([0, 100])

		// The ten vertices, alternating outer and inner. 38.197 and 61.803 are the
		// golden-ratio values the pentagon's diagonals give, which is the arithmetic
		// check that the inner radius came from `adj` and not from a guess.
		expect(result.d).toBe(
			'M 0 38.197 L 38.197 38.197 L 50 0 L 61.803 38.197 L 100 38.197 ' +
				'L 69.098 61.803 L 80.902 100 L 50 76.393 L 19.098 100 L 30.902 61.803 Z'
		)
	})

	it('draws a block arc as two concentric arcs, and does not treat the box as its extent', () => {
		// The default is a half ring across the top: start 180°, end 0°, thickness a
		// quarter of the shorter side. Outer radius 50, inner 25.
		const half = pathOf({ kind: 'preset', preset: 'blockArc', adjustValues: {} }, 100, 100)
		expect(half.d).toBe('M 0 50 A 50 50 0 0 1 100 50 L 75 50 A 25 25 0 0 0 25 50 Z')

		// A quarter turn fills one quadrant and leaves the other three empty — the one
		// preset here whose outline does not reach its own bounding box, which is why
		// it cannot be sanity-checked the way `star5` above is.
		const quarter = pathOf(
			{ kind: 'preset', preset: 'blockArc', adjustValues: { adj1: 'val 0', adj2: 'val 5400000' } },
			100,
			100
		)
		expect(quarter.d).toBe('M 100 50 A 50 50 0 0 1 50 100 L 50 75 A 25 25 0 0 0 75 50 Z')
	})

	it('splits a whole-turn block arc, which as one arc would draw nothing', () => {
		// Equal angles mean the full ring, and SVG derives an arc's centre from its
		// endpoints — so a single 360° `A`, whose endpoints coincide, is degenerate and
		// renders as empty. Two half turns is the same ring and actually draws.
		const ring = pathOf(
			{ kind: 'preset', preset: 'blockArc', adjustValues: { adj1: 'val 0', adj2: 'val 0' } },
			100,
			100
		)
		expect(ring.d).toBe(
			'M 100 50 A 50 50 0 0 1 0 50 A 50 50 0 0 1 100 50 L 75 50 A 25 25 0 0 0 25 50 A 25 25 0 0 0 75 50 Z'
		)
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
