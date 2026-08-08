/**
 * Pure SVG geometry → neutral segment list. The correctness-critical core of
 * SVG-path vectorization, kept **DOM-free** so it is
 * unit-testable in node without a browser (this is where custGeom fidelity lives).
 *
 * It parses an SVG path `d` string (and converts the basic SVG shapes to a `d`)
 * into absolute {@link GeomSeg} commands using only `move` / `line` / `cubic` /
 * `quad` / `close`. Elliptical arcs (`A`/`a`) are approximated as cubic béziers so
 * the output maps one-to-one onto ts-pptx's freeform point DSL — OOXML `arcTo`
 * is center-parameterized and does not round-trip SVG's endpoint arcs cleanly.
 *
 * The browser glue that reads the SVG DOM, resolves colors and scales these
 * segments into slide-space inches lives in `extract/svg.ts`.
 */

/** A neutral, absolute-coordinate path segment in the SVG's own user space. */
export type GeomSeg =
	| { cmd: 'move'; x: number; y: number }
	| { cmd: 'line'; x: number; y: number }
	| { cmd: 'cubic'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
	| { cmd: 'quad'; x1: number; y1: number; x: number; y: number }
	| { cmd: 'close' }

// --- Tokenizer ---------------------------------------------------------------

// Sticky number matcher: optional sign, int/float, optional exponent. Used for
// every numeric path argument *except* arc flags (which are single 0/1 chars and
// may appear unseparated, e.g. `a5 5 0 11 5,5`).
const NUMBER_RE = /[\s,]*([+-]?(?:\d*\.\d+|\d+\.?)(?:[eE][+-]?\d+)?)/y

class Lexer {
	private readonly s: string
	private i = 0
	constructor(s: string) {
		this.s = s
	}

	/** Read the next number, or null at end of meaningful input. */
	number(): number | null {
		NUMBER_RE.lastIndex = this.i
		const m = NUMBER_RE.exec(this.s)
		if (!m || m.index !== this.i) return null
		this.i = NUMBER_RE.lastIndex
		const n = parseFloat(m[1])
		return Number.isFinite(n) ? n : null
	}

	/** Read an arc flag: a single `0` or `1`, after optional separators. */
	flag(): number | null {
		while (this.i < this.s.length && (this.s[this.i] === ' ' || this.s[this.i] === ',' || this.s[this.i] === '\n' || this.s[this.i] === '\t' || this.s[this.i] === '\r')) this.i++
		const c = this.s[this.i]
		if (c !== '0' && c !== '1') return null
		this.i++
		return c === '1' ? 1 : 0
	}

	/** Read the next command letter, or null at end. */
	command(): string | null {
		while (this.i < this.s.length && (this.s[this.i] === ' ' || this.s[this.i] === ',' || this.s[this.i] === '\n' || this.s[this.i] === '\t' || this.s[this.i] === '\r')) this.i++
		const c = this.s[this.i]
		if (c && /[a-zA-Z]/.test(c)) {
			this.i++
			return c
		}
		return null
	}
}

// --- Arc → cubic -------------------------------------------------------------

/**
 * Convert one SVG elliptical-arc segment (endpoint parameterization) into a list
 * of cubic-bézier control sets in absolute coordinates. Splits the sweep into
 * ≤90° pieces for a faithful approximation. Standard implementation per the SVG
 * 1.1 implementation notes (F.6).
 */
function arcToCubics(
	x1: number,
	y1: number,
	rxIn: number,
	ryIn: number,
	phiDeg: number,
	largeArc: number,
	sweep: number,
	x2: number,
	y2: number
): Array<{ x1: number; y1: number; x2: number; y2: number; x: number; y: number }> {
	let rx = Math.abs(rxIn)
	let ry = Math.abs(ryIn)
	// Degenerate radii: SVG spec says treat as a straight line.
	if (rx === 0 || ry === 0 || (x1 === x2 && y1 === y2)) {
		return [{ x1, y1, x2, y2, x: x2, y: y2 }]
	}
	const phi = (phiDeg * Math.PI) / 180
	const cosP = Math.cos(phi)
	const sinP = Math.sin(phi)
	const dx = (x1 - x2) / 2
	const dy = (y1 - y2) / 2
	const x1p = cosP * dx + sinP * dy
	const y1p = -sinP * dx + cosP * dy
	// Scale radii up if they're too small to span the endpoints.
	const lambda = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry)
	if (lambda > 1) {
		const s = Math.sqrt(lambda)
		rx *= s
		ry *= s
	}
	const sign = largeArc !== sweep ? 1 : -1
	const num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p
	const den = rx * rx * y1p * y1p + ry * ry * x1p * x1p
	const co = sign * Math.sqrt(Math.max(0, num / den))
	const cxp = (co * (rx * y1p)) / ry
	const cyp = (co * -(ry * x1p)) / rx
	const cx = cosP * cxp - sinP * cyp + (x1 + x2) / 2
	const cy = sinP * cxp + cosP * cyp + (y1 + y2) / 2

	const angle = (ux: number, uy: number, vx: number, vy: number): number => {
		const dot = ux * vx + uy * vy
		const len = Math.sqrt((ux * ux + uy * uy) * (vx * vx + vy * vy))
		let a = Math.acos(Math.min(1, Math.max(-1, len === 0 ? 1 : dot / len)))
		if (ux * vy - uy * vx < 0) a = -a
		return a
	}
	const theta1 = angle(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry)
	let dTheta = angle((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry)
	if (!sweep && dTheta > 0) dTheta -= 2 * Math.PI
	if (sweep && dTheta < 0) dTheta += 2 * Math.PI

	const segments = Math.max(1, Math.ceil(Math.abs(dTheta) / (Math.PI / 2)))
	const delta = dTheta / segments
	const tau = ((8 / 3) * Math.sin(delta / 4) * Math.sin(delta / 4)) / Math.sin(delta / 2)

	// Point and derivative on the (rotated, translated) ellipse at angle `a`.
	const point = (a: number): [number, number] => [
		cosP * rx * Math.cos(a) - sinP * ry * Math.sin(a) + cx,
		sinP * rx * Math.cos(a) + cosP * ry * Math.sin(a) + cy,
	]
	const deriv = (a: number): [number, number] => [
		-cosP * rx * Math.sin(a) - sinP * ry * Math.cos(a),
		-sinP * rx * Math.sin(a) + cosP * ry * Math.cos(a),
	]

	const out: Array<{ x1: number; y1: number; x2: number; y2: number; x: number; y: number }> = []
	let a1 = theta1
	for (let i = 0; i < segments; i++) {
		const a2 = a1 + delta
		const [p1x, p1y] = point(a1)
		const [d1x, d1y] = deriv(a1)
		const [p2x, p2y] = point(a2)
		const [d2x, d2y] = deriv(a2)
		out.push({
			x1: p1x + tau * d1x,
			y1: p1y + tau * d1y,
			x2: p2x - tau * d2x,
			y2: p2y - tau * d2y,
			x: p2x,
			y: p2y,
		})
		a1 = a2
	}
	return out
}

// --- Path `d` parser ---------------------------------------------------------

/**
 * Parse an SVG path `d` string into absolute {@link GeomSeg} commands. Returns an
 * empty array for empty/whitespace input. Throws nothing: malformed tails stop
 * parsing (best-effort), so the caller can fall back to raster if the result is
 * unusable.
 */
export function parsePathData(d: string): GeomSeg[] {
	const segs: GeomSeg[] = []
	if (!d || !/[a-zA-Z]/.test(d)) return segs
	const lex = new Lexer(d)
	let cx = 0
	let cy = 0
	let startX = 0
	let startY = 0
	// Previous cubic/quad control point (absolute), for S/T reflection.
	let prevCubicCtrl: [number, number] | null = null
	let prevQuadCtrl: [number, number] | null = null
	let cmd = lex.command()

	while (cmd) {
		const rel = cmd === cmd.toLowerCase()
		const C = cmd.toUpperCase()
		const ox = rel ? cx : 0
		const oy = rel ? cy : 0

		switch (C) {
			case 'M': {
				const x = lex.number()
				const y = lex.number()
				if (x === null || y === null) return segs
				cx = ox + x
				cy = oy + y
				startX = cx
				startY = cy
				segs.push({ cmd: 'move', x: cx, y: cy })
				prevCubicCtrl = prevQuadCtrl = null
				// Subsequent implicit pairs after M are treated as L (per spec).
				for (;;) {
					const nx = lex.number()
					if (nx === null) break
					const ny = lex.number()
					if (ny === null) return segs
					cx = (rel ? cx : 0) + nx
					cy = (rel ? cy : 0) + ny
					segs.push({ cmd: 'line', x: cx, y: cy })
				}
				break
			}
			case 'L': {
				for (;;) {
					const x = lex.number()
					if (x === null) break
					const y = lex.number()
					if (y === null) return segs
					cx = (rel ? cx : 0) + x
					cy = (rel ? cy : 0) + y
					segs.push({ cmd: 'line', x: cx, y: cy })
				}
				prevCubicCtrl = prevQuadCtrl = null
				break
			}
			case 'H': {
				for (;;) {
					const x = lex.number()
					if (x === null) break
					cx = (rel ? cx : 0) + x
					segs.push({ cmd: 'line', x: cx, y: cy })
				}
				prevCubicCtrl = prevQuadCtrl = null
				break
			}
			case 'V': {
				for (;;) {
					const y = lex.number()
					if (y === null) break
					cy = (rel ? cy : 0) + y
					segs.push({ cmd: 'line', x: cx, y: cy })
				}
				prevCubicCtrl = prevQuadCtrl = null
				break
			}
			case 'C': {
				for (;;) {
					const x1 = lex.number()
					if (x1 === null) break
					const y1 = lex.number()
					const x2 = lex.number()
					const y2 = lex.number()
					const x = lex.number()
					const y = lex.number()
					if (y1 === null || x2 === null || y2 === null || x === null || y === null) return segs
					const ax1 = ox + x1
					const ay1 = oy + y1
					const ax2 = ox + x2
					const ay2 = oy + y2
					const axx = ox + x
					const ayy = oy + y
					segs.push({ cmd: 'cubic', x1: ax1, y1: ay1, x2: ax2, y2: ay2, x: axx, y: ayy })
					cx = axx
					cy = ayy
					prevCubicCtrl = [ax2, ay2]
					prevQuadCtrl = null
				}
				break
			}
			case 'S': {
				for (;;) {
					const x2 = lex.number()
					if (x2 === null) break
					const y2 = lex.number()
					const x = lex.number()
					const y = lex.number()
					if (y2 === null || x === null || y === null) return segs
					const rx1 = prevCubicCtrl ? 2 * cx - prevCubicCtrl[0] : cx
					const ry1 = prevCubicCtrl ? 2 * cy - prevCubicCtrl[1] : cy
					const ax2 = ox + x2
					const ay2 = oy + y2
					const axx = ox + x
					const ayy = oy + y
					segs.push({ cmd: 'cubic', x1: rx1, y1: ry1, x2: ax2, y2: ay2, x: axx, y: ayy })
					cx = axx
					cy = ayy
					prevCubicCtrl = [ax2, ay2]
					prevQuadCtrl = null
				}
				break
			}
			case 'Q': {
				for (;;) {
					const x1 = lex.number()
					if (x1 === null) break
					const y1 = lex.number()
					const x = lex.number()
					const y = lex.number()
					if (y1 === null || x === null || y === null) return segs
					const ax1 = ox + x1
					const ay1 = oy + y1
					const axx = ox + x
					const ayy = oy + y
					segs.push({ cmd: 'quad', x1: ax1, y1: ay1, x: axx, y: ayy })
					cx = axx
					cy = ayy
					prevQuadCtrl = [ax1, ay1]
					prevCubicCtrl = null
				}
				break
			}
			case 'T': {
				for (;;) {
					const x = lex.number()
					if (x === null) break
					const y = lex.number()
					if (y === null) return segs
					const qx: number = prevQuadCtrl ? 2 * cx - prevQuadCtrl[0] : cx
					const qy: number = prevQuadCtrl ? 2 * cy - prevQuadCtrl[1] : cy
					const axx = ox + x
					const ayy = oy + y
					segs.push({ cmd: 'quad', x1: qx, y1: qy, x: axx, y: ayy })
					cx = axx
					cy = ayy
					prevQuadCtrl = [qx, qy]
					prevCubicCtrl = null
				}
				break
			}
			case 'A': {
				for (;;) {
					const rx = lex.number()
					if (rx === null) break
					const ry = lex.number()
					const rot = lex.number()
					const large = lex.flag()
					const sweep = lex.flag()
					const x = lex.number()
					const y = lex.number()
					if (ry === null || rot === null || large === null || sweep === null || x === null || y === null) return segs
					const axx = ox + x
					const ayy = oy + y
					for (const c of arcToCubics(cx, cy, rx, ry, rot, large, sweep, axx, ayy)) {
						segs.push({ cmd: 'cubic', x1: c.x1, y1: c.y1, x2: c.x2, y2: c.y2, x: c.x, y: c.y })
					}
					cx = axx
					cy = ayy
					prevCubicCtrl = prevQuadCtrl = null
				}
				break
			}
			case 'Z': {
				segs.push({ cmd: 'close' })
				cx = startX
				cy = startY
				prevCubicCtrl = prevQuadCtrl = null
				break
			}
			default:
				// Unknown command: stop best-effort parsing.
				return segs
		}
		cmd = lex.command()
	}
	return segs
}

// --- Basic shape → `d` -------------------------------------------------------

const num = (v: string | null | undefined): number => {
	const n = parseFloat(String(v ?? ''))
	return Number.isFinite(n) ? n : 0
}

/** `<rect>` → `d`. Supports rounded corners (`rx`/`ry`) via arcs. */
export function rectToPathData(x: string | null, y: string | null, w: string | null, h: string | null, rx?: string | null, ry?: string | null): string {
	const X = num(x)
	const Y = num(y)
	const W = num(w)
	const H = num(h)
	if (W <= 0 || H <= 0) return ''
	let RX = rx == null || rx === '' ? NaN : num(rx)
	let RY = ry == null || ry === '' ? NaN : num(ry)
	if (Number.isNaN(RX) && !Number.isNaN(RY)) RX = RY
	if (Number.isNaN(RY) && !Number.isNaN(RX)) RY = RX
	if (Number.isNaN(RX)) RX = 0
	if (Number.isNaN(RY)) RY = 0
	RX = Math.min(Math.max(0, RX), W / 2)
	RY = Math.min(Math.max(0, RY), H / 2)
	if (RX === 0 || RY === 0) {
		return `M${X},${Y} H${X + W} V${Y + H} H${X} Z`
	}
	return (
		`M${X + RX},${Y} ` +
		`H${X + W - RX} A${RX},${RY} 0 0 1 ${X + W},${Y + RY} ` +
		`V${Y + H - RY} A${RX},${RY} 0 0 1 ${X + W - RX},${Y + H} ` +
		`H${X + RX} A${RX},${RY} 0 0 1 ${X},${Y + H - RY} ` +
		`V${Y + RY} A${RX},${RY} 0 0 1 ${X + RX},${Y} Z`
	)
}

/** `<circle>` → `d` (two semicircle arcs). */
export function circleToPathData(cx: string | null, cy: string | null, r: string | null): string {
	return ellipseToPathData(cx, cy, r, r)
}

/** `<ellipse>` → `d` (two half arcs). */
export function ellipseToPathData(cx: string | null, cy: string | null, rx: string | null, ry: string | null): string {
	const CX = num(cx)
	const CY = num(cy)
	const RX = num(rx)
	const RY = num(ry)
	if (RX <= 0 || RY <= 0) return ''
	return `M${CX - RX},${CY} A${RX},${RY} 0 1 0 ${CX + RX},${CY} A${RX},${RY} 0 1 0 ${CX - RX},${CY} Z`
}

/** `<line>` → `d`. */
export function lineToPathData(x1: string | null, y1: string | null, x2: string | null, y2: string | null): string {
	return `M${num(x1)},${num(y1)} L${num(x2)},${num(y2)}`
}

/** `<polygon>`/`<polyline>` points → `d`. `close` adds the trailing `Z`. */
export function pointsToPathData(points: string | null, close: boolean): string {
	const nums = String(points ?? '')
		.trim()
		.split(/[\s,]+/)
		.map(Number)
		.filter((n) => Number.isFinite(n))
	if (nums.length < 4) return ''
	let d = `M${nums[0]},${nums[1]}`
	for (let i = 2; i + 1 < nums.length; i += 2) d += ` L${nums[i]},${nums[i + 1]}`
	if (close) d += ' Z'
	return d
}
