/**
 * Pure SVG-geometry parser, headless. Exercises `parsePathData` and the
 * basic-shape → `d` helpers. These run in node with no DOM: the
 * correctness-critical core of vectorization lives here, where it's cheap to test
 * exhaustively. The DOM glue (`heuristic/svg.ts`) is covered by the browser layer.
 */

import { describe, expect, it } from 'vitest'
import {
	circleToPathData,
	ellipseToPathData,
	type GeomSeg,
	lineToPathData,
	parsePathData,
	pointsToPathData,
	rectToPathData,
} from '../../src/heuristic/svg-path'

describe('parsePathData — commands', () => {
	it('parses an absolute move + line triangle with close', () => {
		expect(parsePathData('M0 0 L10 0 L10 10 Z')).toEqual<GeomSeg[]>([
			{ cmd: 'move', x: 0, y: 0 },
			{ cmd: 'line', x: 10, y: 0 },
			{ cmd: 'line', x: 10, y: 10 },
			{ cmd: 'close' },
		])
	})

	it('resolves relative commands against the current point', () => {
		// m moves to (1,1); l is relative → (1+2, 1+3) = (3,4); h/v relative.
		expect(parsePathData('m1 1 l2 3 h2 v-1')).toEqual<GeomSeg[]>([
			{ cmd: 'move', x: 1, y: 1 },
			{ cmd: 'line', x: 3, y: 4 },
			{ cmd: 'line', x: 5, y: 4 },
			{ cmd: 'line', x: 5, y: 3 },
		])
	})

	it('treats extra coordinate pairs after M as implicit lineTo', () => {
		expect(parsePathData('M1 2 3 4 5 6')).toEqual<GeomSeg[]>([
			{ cmd: 'move', x: 1, y: 2 },
			{ cmd: 'line', x: 3, y: 4 },
			{ cmd: 'line', x: 5, y: 6 },
		])
	})

	it('parses cubic and reflects S off the previous control point', () => {
		const segs = parsePathData('M0 0 C1 1 2 1 3 0 S5 -1 6 0')
		expect(segs[0]).toEqual({ cmd: 'move', x: 0, y: 0 })
		expect(segs[1]).toEqual({ cmd: 'cubic', x1: 1, y1: 1, x2: 2, y2: 1, x: 3, y: 0 })
		// Reflection of (2,1) about (3,0) → (4,-1).
		expect(segs[2]).toEqual({ cmd: 'cubic', x1: 4, y1: -1, x2: 5, y2: -1, x: 6, y: 0 })
	})

	it('parses quadratic and reflects T off the previous control point', () => {
		const segs = parsePathData('M0 0 Q1 2 2 0 T4 0')
		expect(segs[1]).toEqual({ cmd: 'quad', x1: 1, y1: 2, x: 2, y: 0 })
		// Reflection of (1,2) about (2,0) → (3,-2).
		expect(segs[2]).toEqual({ cmd: 'quad', x1: 3, y1: -2, x: 4, y: 0 })
	})

	it('handles tight number packing and exponents', () => {
		expect(parsePathData('M.5.5L-1-1')).toEqual<GeomSeg[]>([
			{ cmd: 'move', x: 0.5, y: 0.5 },
			{ cmd: 'line', x: -1, y: -1 },
		])
		expect(parsePathData('M1e1 2e1')).toEqual<GeomSeg[]>([{ cmd: 'move', x: 10, y: 20 }])
	})

	it('returns empty for empty / non-path input', () => {
		expect(parsePathData('')).toEqual([])
		expect(parsePathData('   ')).toEqual([])
	})
})

describe('parsePathData — arcs become cubics', () => {
	it('converts an A arc into one or more cubic segments ending at the arc endpoint', () => {
		const segs = parsePathData('M0 0 A5 5 0 0 1 10 0')
		expect(segs[0]).toEqual({ cmd: 'move', x: 0, y: 0 })
		expect(segs.length).toBeGreaterThan(1)
		expect(segs.every((s, i) => i === 0 || s.cmd === 'cubic')).toBe(true)
		const last = segs[segs.length - 1] as Extract<GeomSeg, { cmd: 'cubic' }>
		expect(last.x).toBeCloseTo(10, 4)
		expect(last.y).toBeCloseTo(0, 4)
	})

	it('approximates a semicircle arc with its midpoint bulging out to the radius', () => {
		// Half-circle radius 5 from (0,0) to (10,0): split into two 90° cubics whose
		// shared endpoint (the arc midpoint) sits a full radius off the chord.
		const segs = parsePathData('M0 0 A5 5 0 0 1 10 0') as Array<Extract<GeomSeg, { cmd: 'cubic' }>>
		const peak = Math.max(...segs.slice(1).map((s) => Math.abs(s.y)))
		expect(peak).toBeCloseTo(5, 4)
	})

	it('degenerate (zero-radius) arc collapses to a line to the endpoint', () => {
		const segs = parsePathData('M0 0 A0 0 0 0 1 10 5')
		const last = segs[segs.length - 1] as Extract<GeomSeg, { cmd: 'cubic' }>
		expect(last.x).toBeCloseTo(10, 6)
		expect(last.y).toBeCloseTo(5, 6)
	})
})

describe('basic shape → d', () => {
	it('rect (sharp) is a closed rectangle', () => {
		expect(parsePathData(rectToPathData('1', '2', '4', '3'))).toEqual<GeomSeg[]>([
			{ cmd: 'move', x: 1, y: 2 },
			{ cmd: 'line', x: 5, y: 2 },
			{ cmd: 'line', x: 5, y: 5 },
			{ cmd: 'line', x: 1, y: 5 },
			{ cmd: 'close' },
		])
	})

	it('rounded rect uses arcs (→ cubics) and clamps radius to half the side', () => {
		const segs = parsePathData(rectToPathData('0', '0', '10', '10', '100', '100'))
		// rx/ry clamp to 5; corners become cubic segments.
		expect(segs.some((s) => s.cmd === 'cubic')).toBe(true)
		expect(segs[segs.length - 1]).toEqual({ cmd: 'close' })
	})

	it('circle spans its diameter horizontally', () => {
		const segs = parsePathData(circleToPathData('5', '5', '5')) as GeomSeg[]
		const move = segs[0] as Extract<GeomSeg, { cmd: 'move' }>
		expect(move.x).toBeCloseTo(0, 6)
		expect(move.y).toBeCloseTo(5, 6)
	})

	it('ellipse with zero radius yields no path', () => {
		expect(ellipseToPathData('5', '5', '0', '3')).toBe('')
	})

	it('line is a single move + line', () => {
		expect(parsePathData(lineToPathData('0', '0', '4', '2'))).toEqual<GeomSeg[]>([
			{ cmd: 'move', x: 0, y: 0 },
			{ cmd: 'line', x: 4, y: 2 },
		])
	})

	it('polygon closes; polyline does not', () => {
		expect(parsePathData(pointsToPathData('0,0 4,0 4,4', true))).toEqual<GeomSeg[]>([
			{ cmd: 'move', x: 0, y: 0 },
			{ cmd: 'line', x: 4, y: 0 },
			{ cmd: 'line', x: 4, y: 4 },
			{ cmd: 'close' },
		])
		const open = parsePathData(pointsToPathData('0,0 4,0 4,4', false))
		expect(open[open.length - 1]).toEqual({ cmd: 'line', x: 4, y: 4 })
	})
})
