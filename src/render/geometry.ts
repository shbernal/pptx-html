/**
 * {@link Geometry} → SVG path data.
 *
 * Everything here works in a node's **local** space, `0..w` × `0..h` EMU, and
 * `document.ts` places the result with one `translate`/`rotate`/`scale` group.
 * Drawing in local space is what makes rotation and flips a single transform
 * around `(w/2, h/2)` instead of arithmetic repeated in every path command.
 *
 * ## Presets are resolved here, and that is not a second source of truth
 *
 * There is no upstream preset catalogue to defer to. `@shbernal/ts-pptx` emits
 * `<a:prstGeom prst="…">` by name and lets PowerPoint resolve it —
 * `VALID_SHAPE_PRESETS` is 188 names and no geometry — and the read model's
 * `customGeometry` is documented as `null` for exactly the shapes that use a
 * preset. The full formulas live in ECMA-376 Annex D's electronic addenda, which
 * neither library ships.
 *
 * Each entry below is therefore transcribed from that addendum's
 * `presetShapeDefinitions.xml` and reduced by hand, with the guide operators read
 * off ECMA-376 §20.1.9.11 — the multiply-divide operator is `(x * y) / z`, `pin`
 * clamps its middle argument between the outer two, and `?:` branches on whether
 * its first is positive — and the built-in guides off §20.1.10.55, where `ss` is
 * the shorter side and `hc`/`vc` the centre. The reduction is checked by the fact
 * that `parallelogram`'s definition collapses to the one line it already had.
 *
 * So this is a **picture**, and it is allowed to be an incomplete one. The IR
 * carries `preset` and `adjustValues` verbatim through the island, so a shape
 * this file cannot draw still round-trips byte-for-byte — it just looks like a
 * rectangle in the preview. That is the difference between the two channels
 * doing their jobs and a silent loss, and it is why the unresolved case is
 * *marked* ({@link PathResult.fallback}) rather than quietly boxed.
 */

import { ANGLE_UNITS_PER_DEGREE } from '@shbernal/ts-pptx'
import type { Geometry, GeometryCommand, GeometryPath } from '../ir/render'
import { round3 } from './paint'

export interface PathResult {
	/** SVG path data in local EMU space. */
	d: string
	/**
	 * The preset name, when this is a bounding box standing in for a preset with
	 * no local formula. `document.ts` writes it into the DOM so an approximated
	 * outline is visible as such rather than passing for the real thing.
	 */
	fallback?: string
	/** `@fill` of each source path, when the geometry states them (`none` suppresses fill). */
	unfilled?: boolean
	/** `@stroke="false"` on every source path. */
	unstroked?: boolean
}

/**
 * `a:avLst` guide values are a formula language, not numbers — the IR carries
 * them as the raw `a:gd/@fmla` strings for that reason. Only the `val N` form
 * states a literal, and it is the only form PowerPoint itself writes into an
 * `a:avLst`; a computed guide there would mean the shape's adjust handle is
 * driven by another guide, which no preset below models.
 *
 * Anything else falls back to the preset's own default, which is the same thing
 * that happens when the attribute is absent.
 */
function adjust(values: Record<string, string>, name: string, fallback: number): number {
	const formula = values[name]
	if (formula === undefined) return fallback
	const match = /^val\s+(-?\d+)$/.exec(formula.trim())
	return match?.[1] === undefined ? fallback : Number(match[1])
}

/** Guide values are hundred-thousandths, so `16667` is 16.667 %. */
const GUIDE_SCALE = 100_000

// OOXML states angles in 60,000ths of a degree, so a right angle is `5400000`.
// `ANGLE_UNITS_PER_DEGREE` is upstream's name for that number: an internal-only
// conversion, so it comes from the library that owns the format rather than
// being written down a second time here. See `src/constants.ts`.
function radians(angle: number): number {
	return ((angle / ANGLE_UNITS_PER_DEGREE) * Math.PI) / 180
}

/** The `pin x y z` guide formula: `y`, held inside `[x, z]`. */
function pin(low: number, value: number, high: number): number {
	return Math.min(Math.max(value, low), high)
}

/**
 * A point on an ellipse at an OOXML *nominal* angle, relative to its centre.
 *
 * The `cat2`/`sat2` pair in a preset definition, which is not the same as
 * `(rx cos θ, ry sin θ)`: OOXML's angle is the direction of a ray from the
 * centre, and the point wanted is where that ray meets the ellipse, whose
 * parametric angle differs unless the ellipse is a circle.
 *
 * ECMA-376 §20.1.9.11 writes both as `arctan(z / y)`, a single-argument arctan
 * that cannot tell one diagonal from its opposite — under it a `stAng` of 180°
 * yields `arctan(-0)`, so the point comes out on the *right* of the shape rather
 * than the left. `atan2` is what makes the quadrant survive, and what every
 * renderer that draws these shapes correctly uses.
 */
function onEllipse(rx: number, ry: number, angle: number): { x: number; y: number } {
	const parametric = Math.atan2(rx * Math.sin(radians(angle)), ry * Math.cos(radians(angle)))
	return { x: rx * Math.cos(parametric), y: ry * Math.sin(parametric) }
}

/**
 * Trig lands values like `61.80339887498949` where the surrounding presets, whose
 * formulas are arithmetic on EMU, land exact ones. Everything here is EMU — a
 * 914,400th of an inch — so three decimals is already far below anything a screen
 * can show, and the alternative is path data whose noise obscures the shape.
 *
 * Applied only to the trig presets, so the exact ones keep printing exactly.
 */
function coord(value: number): number {
	return round3(value)
}

/** A full turn, in OOXML's angle units. Half of it is the SVG large-arc threshold. */
const FULL_TURN = 21_600_000

/**
 * One elliptical arc as SVG `A` commands, given a start angle and a signed sweep.
 *
 * Two commands rather than one when the sweep is a whole turn, because SVG derives
 * an arc's centre from its *endpoints* — so a 360° arc, whose endpoints coincide,
 * is degenerate and draws nothing at all. A `blockArc` whose two angles are equal
 * is exactly that case, and it means the full ring rather than an empty one, so
 * without the split the shape silently disappears instead of looking wrong.
 */
function arcTo(
	rx: number,
	ry: number,
	from: number,
	sweep: number,
	at: (rx: number, ry: number, angle: number) => { x: number; y: number }
): string {
	// A positive sweep is clockwise in both systems, so it is SVG's sweep flag 1
	// with no sign juggling — the same reasoning as `segmentOf`.
	const flag = sweep >= 0 ? 1 : 0
	// The large-arc flag is a property of one step, not of the list, so it is read
	// off the value the list was built from. Both halves of a split sweep are the
	// same size, which is why one `step` answers for either shape of `steps`.
	const split = Math.abs(sweep) >= FULL_TURN
	const step = split ? sweep / 2 : sweep
	const steps = split ? [step, step] : [step]
	const large = Math.abs(step) > FULL_TURN / 2 ? 1 : 0

	let angle = from
	return steps
		.map((each) => {
			angle += each
			const end = at(rx, ry, angle)
			return `A ${coord(rx)} ${coord(ry)} 0 ${large} ${flag} ${end.x} ${end.y}`
		})
		.join(' ')
}

function rect(w: number, h: number): string {
	return `M 0 0 L ${w} 0 L ${w} ${h} L 0 ${h} Z`
}

/**
 * The presets with a formula simple enough to state exactly, which between them
 * cover the great majority of shapes on a real slide. Each takes the box and the
 * source `a:avLst`; anything not here draws its bounding box and says so.
 *
 * Deliberately not a race to 188: each entry is a claim that this file draws the
 * shape *correctly*, and a wrong outline is worse than an obvious box, because
 * only one of the two looks like an error.
 *
 * Which is also why additions are chosen by counting rather than by taste.
 * Across 47 PowerPoint-authored decks the original nine resolved 670 of 682
 * preset instances; the whole remainder was `chevron` ×9, `star5` ×2 and
 * `blockArc` ×1. A later census of slide-ui's 98 generated browser previews
 * supplied the next evidence set: `round2DiagRect` ×6, `rightArrowCallout` ×4,
 * `rightArrow` ×2, `round2SameRect` ×2 and `wedgeRectCallout` ×2. Entries outside
 * those measured sets still fall back visibly instead of turning an appealing
 * but unevidenced list into dozens of formulas that each have to be right.
 */
/** The four corner radii of a rounded rectangle, clockwise from the top left. */
interface Corners {
	tl: number
	tr: number
	br: number
	bl: number
}

/**
 * A rectangle with four independently rounded corners.
 *
 * One arc sequence rather than three transcriptions of it. The three presets
 * below differ only in which radius each corner gets, and the file's own bar for
 * an entry in `PRESETS` is that it draws the shape *correctly* — three copies of
 * `M / L / A x4 / Z` is three chances to put a sweep flag the wrong way round,
 * and the result of that is an outline subtly wrong rather than obviously a box.
 *
 * Every arc is `0 0 1`: no rotation, small arc, clockwise, which is what a corner
 * traversed clockwise from the top-left always is.
 */
function roundedRectPath(w: number, h: number, corners: Corners): string {
	const { tl, tr, br, bl } = corners
	return [
		`M ${tl} 0`,
		`L ${w - tr} 0`,
		`A ${tr} ${tr} 0 0 1 ${w} ${tr}`,
		`L ${w} ${h - br}`,
		`A ${br} ${br} 0 0 1 ${w - br} ${h}`,
		`L ${bl} ${h}`,
		`A ${bl} ${bl} 0 0 1 0 ${h - bl}`,
		`L 0 ${tl}`,
		`A ${tl} ${tl} 0 0 1 ${tl} 0`,
		'Z',
	].join(' ')
}

const PRESETS: Record<string, (w: number, h: number, adj: Record<string, string>) => string> = {
	rect,
	// A picture's default geometry, and the shape every unresolved preset becomes.
	roundRect: (w, h, adj) => {
		const radius = (adjust(adj, 'adj', 16_667) / GUIDE_SCALE) * Math.min(w, h)
		const r = Math.max(0, Math.min(radius, Math.min(w, h) / 2))
		return roundedRectPath(w, h, { tl: r, tr: r, br: r, bl: r })
	},
	ellipse: (w, h) => {
		const rx = w / 2
		const ry = h / 2
		// Two half-arcs rather than `<ellipse>`, so every geometry in this file is
		// one `<path>` and the caller never branches on element name.
		return `M 0 ${ry} A ${rx} ${ry} 0 1 1 ${w} ${ry} A ${rx} ${ry} 0 1 1 0 ${ry} Z`
	},
	triangle: (w, h, adj) => {
		const apex = (adjust(adj, 'adj', 50_000) / GUIDE_SCALE) * w
		return `M ${apex} 0 L ${w} ${h} L 0 ${h} Z`
	},
	rtTriangle: (w, h) => `M 0 0 L 0 ${h} L ${w} ${h} Z`,
	diamond: (w, h) => `M ${w / 2} 0 L ${w} ${h / 2} L ${w / 2} ${h} L 0 ${h / 2} Z`,
	// `line` and `straightConnector1` are the same outline: the box's diagonal,
	// which flips carry into the other three directions.
	line: (w, h) => `M 0 0 L ${w} ${h}`,
	straightConnector1: (w, h) => `M 0 0 L ${w} ${h}`,
	parallelogram: (w, h, adj) => {
		const inset = Math.min((adjust(adj, 'adj', 25_000) / GUIDE_SCALE) * Math.min(w, h), w)
		return `M ${inset} 0 L ${w} 0 L ${w - inset} ${h} L 0 ${h} Z`
	},
	round2DiagRect: (w, h, adj) => {
		// `adj1` rounds top-left + bottom-right; `adj2` rounds the opposite
		// diagonal. The default second radius is zero, so those two corners remain
		// square unless the deck states otherwise.
		const r1 = (pin(0, adjust(adj, 'adj1', 16_667), 50_000) / GUIDE_SCALE) * Math.min(w, h)
		const r2 = (pin(0, adjust(adj, 'adj2', 0), 50_000) / GUIDE_SCALE) * Math.min(w, h)
		return roundedRectPath(w, h, { tl: r1, tr: r2, br: r1, bl: r2 })
	},
	round2SameRect: (w, h, adj) => {
		// One handle controls both top corners and the other both bottom corners.
		// PowerPoint's default is the useful "rounded header" case: rounded above,
		// square below. Stated as four corners for the first time here, which is what
		// makes the pairing checkable: `tl`/`tr` take `adj1` and `bl`/`br` take
		// `adj2`, and the earlier hand-written path did the same thing less legibly.
		const top = (pin(0, adjust(adj, 'adj1', 16_667), 50_000) / GUIDE_SCALE) * Math.min(w, h)
		const bottom = (pin(0, adjust(adj, 'adj2', 0), 50_000) / GUIDE_SCALE) * Math.min(w, h)
		return roundedRectPath(w, h, { tl: top, tr: top, br: bottom, bl: bottom })
	},
	// The notch and the point are the same inset, which is why one adjust states
	// both. `maxAdj` pins it to `w`, so the two never cross into a bowtie.
	chevron: (w, h, adj) => {
		const x1 = Math.min((adjust(adj, 'adj', 50_000) / GUIDE_SCALE) * Math.min(w, h), w)
		return `M 0 0 L ${w - x1} 0 L ${w} ${h / 2} L ${w - x1} ${h} L 0 ${h} L ${x1} ${h / 2} Z`
	},
	rightArrow: (w, h, adj) => {
		const shaft = pin(0, adjust(adj, 'adj1', 50_000), 100_000)
		const maxHead = Math.min(w, h) === 0 ? 0 : (GUIDE_SCALE * w) / Math.min(w, h)
		const head = pin(0, adjust(adj, 'adj2', 50_000), maxHead)
		const x1 = w - (Math.min(w, h) * head) / GUIDE_SCALE
		const dy = (h * shaft) / (2 * GUIDE_SCALE)
		return `M 0 ${h / 2 - dy} L ${x1} ${h / 2 - dy} L ${x1} 0 L ${w} ${h / 2} L ${x1} ${h} L ${x1} ${h / 2 + dy} L 0 ${h / 2 + dy} Z`
	},
	rightArrowCallout: (w, h, adj) => {
		const ss = Math.min(w, h)
		const maxAdj2 = ss === 0 ? 0 : (50_000 * h) / ss
		const a2 = pin(0, adjust(adj, 'adj2', 25_000), maxAdj2)
		const a1 = pin(0, adjust(adj, 'adj1', 25_000), 2 * a2)
		const maxAdj3 = ss === 0 ? 0 : (GUIDE_SCALE * w) / ss
		const a3 = pin(0, adjust(adj, 'adj3', 25_000), maxAdj3)
		const maxAdj4 = w === 0 ? 0 : GUIDE_SCALE - (a3 * ss) / w
		const a4 = pin(0, adjust(adj, 'adj4', 64_977), maxAdj4)
		const dy1 = (ss * a2) / GUIDE_SCALE
		const dy2 = (ss * a1) / (2 * GUIDE_SCALE)
		const [y1, y2, y3, y4] = [h / 2 - dy1, h / 2 - dy2, h / 2 + dy2, h / 2 + dy1]
		const x3 = w - (ss * a3) / GUIDE_SCALE
		const x2 = (w * a4) / GUIDE_SCALE
		return [
			'M 0 0',
			`L ${x2} 0`,
			`L ${x2} ${y2}`,
			`L ${x3} ${y2}`,
			`L ${x3} ${y1}`,
			`L ${w} ${h / 2}`,
			`L ${x3} ${y4}`,
			`L ${x3} ${y3}`,
			`L ${x2} ${y3}`,
			`L ${x2} ${h}`,
			`L 0 ${h}`,
			'Z',
		].join(' ')
	},
	wedgeRectCallout: (w, h, adj) => {
		const dxPos = (w * adjust(adj, 'adj1', -20_833)) / GUIDE_SCALE
		const dyPos = (h * adjust(adj, 'adj2', 62_500)) / GUIDE_SCALE
		const xPos = w / 2 + dxPos
		const yPos = h / 2 + dyPos
		const dq = w === 0 ? 0 : (dxPos * h) / w
		const dz = Math.abs(dyPos) - Math.abs(dq)
		const [x1, x2] = dxPos > 0 ? [(w * 7) / 12, (w * 10) / 12] : [(w * 2) / 12, (w * 5) / 12]
		const [y1, y2] = dyPos > 0 ? [(h * 7) / 12, (h * 10) / 12] : [(h * 2) / 12, (h * 5) / 12]

		// DR-18-0013 corrects four edge guides in the published electronic
		// addendum. Without these clamps a handle dragged *inside* the rectangle
		// cuts a notch into it; PowerPoint instead collapses that wedge to the edge.
		const xl = Math.min(dz > 0 ? 0 : dxPos > 0 ? 0 : xPos, 0)
		const xt = dz > 0 ? (dyPos > 0 ? x1 : xPos) : x1
		const xr = Math.max(dz > 0 ? w : dxPos > 0 ? xPos : w, w)
		const xb = dz > 0 ? (dyPos > 0 ? xPos : x1) : x1
		const yl = dz > 0 ? y1 : dxPos > 0 ? y1 : yPos
		const yt = Math.min(dz > 0 ? (dyPos > 0 ? 0 : yPos) : 0, 0)
		const yr = dz > 0 ? y1 : dxPos > 0 ? yPos : y1
		const yb = Math.max(dz > 0 ? (dyPos > 0 ? yPos : h) : h, h)

		return [
			'M 0 0',
			`L ${x1} 0`,
			`L ${xt} ${yt}`,
			`L ${x2} 0`,
			`L ${w} 0`,
			`L ${w} ${y1}`,
			`L ${xr} ${yr}`,
			`L ${w} ${y2}`,
			`L ${w} ${h}`,
			`L ${x2} ${h}`,
			`L ${xb} ${yb}`,
			`L ${x1} ${h}`,
			`L 0 ${h}`,
			`L 0 ${y2}`,
			`L ${xl} ${yl}`,
			`L 0 ${y1}`,
			'Z',
		].join(' ')
	},
	star5: (w, h, adj) => {
		// `hf`/`vf` are in `a:avLst` beside `adj`, so a deck may state them — but
		// their defaults are not arbitrary: they scale the pentagon's circumscribed
		// circle so the five outer points land exactly on the box's edges. Without
		// them the star would sit inset with the box's corners empty.
		const a = pin(0, adjust(adj, 'adj', 19_098), 50_000)
		const swd2 = (w / 2) * (adjust(adj, 'hf', 105_146) / GUIDE_SCALE)
		const vf = adjust(adj, 'vf', 110_557) / GUIDE_SCALE
		const shd2 = (h / 2) * vf
		// The vertical centre moves with `vf` too, which is what keeps the top point
		// on `y = 0` rather than above it.
		const svc = (h / 2) * vf
		const hc = w / 2

		const outer = (angle: number) => ({ x: swd2 * Math.cos(radians(angle)), y: shd2 * Math.sin(radians(angle)) })
		// 18° and 306°: the two pairs of outer points either side of the apex.
		const d1 = outer(1_080_000)
		const d2 = outer(18_360_000)
		const [x1, x2, x3, x4] = [hc - d1.x, hc - d2.x, hc + d2.x, hc + d1.x]
		const [y1, y2] = [svc - d1.y, svc - d2.y]

		// The inner pentagon, at `adj` of the outer radius — the notch depth.
		const iwd2 = swd2 * (a / 50_000)
		const ihd2 = shd2 * (a / 50_000)
		const inner = (angle: number) => ({ x: iwd2 * Math.cos(radians(angle)), y: ihd2 * Math.sin(radians(angle)) })
		const s1 = inner(20_520_000)
		const s2 = inner(3_240_000)
		const [sx1, sx2, sx3, sx4] = [hc - s1.x, hc - s2.x, hc + s2.x, hc + s1.x]
		const [sy1, sy2, sy3] = [svc - s2.y, svc - s1.y, svc + ihd2]

		return [
			`M ${coord(x1)} ${coord(y1)}`,
			`L ${coord(sx2)} ${coord(sy1)}`,
			`L ${coord(hc)} 0`,
			`L ${coord(sx3)} ${coord(sy1)}`,
			`L ${coord(x4)} ${coord(y1)}`,
			`L ${coord(sx4)} ${coord(sy2)}`,
			`L ${coord(x3)} ${coord(y2)}`,
			`L ${coord(hc)} ${coord(sy3)}`,
			`L ${coord(x2)} ${coord(y2)}`,
			`L ${coord(sx1)} ${coord(sy2)}`,
			'Z',
		].join(' ')
	},
	blockArc: (w, h, adj) => {
		// Two concentric elliptical arcs joined at both ends. `adj1`/`adj2` are the
		// start and end angles and `adj3` is the ring's thickness, so unlike every
		// other preset here the outline is not a polygon and the box is not its
		// extent — a 90° block arc occupies one quadrant and leaves the rest empty.
		const stAng = pin(0, adjust(adj, 'adj1', 10_800_000), 21_599_999)
		const istAng = pin(0, adjust(adj, 'adj2', 0), 21_599_999)
		const thickness = pin(0, adjust(adj, 'adj3', 25_000), 50_000)

		// A sweep that would be zero or negative is the long way round instead: equal
		// angles mean the whole ring, not an empty one.
		const gap = istAng - stAng
		const swAng = gap > 0 ? gap : gap + 21_600_000

		const [wd2, hd2] = [w / 2, h / 2]
		// Pinned to 50000, so `dr` never exceeds half the shorter side and the inner
		// radii never go negative. At the maximum they reach zero and this is a wedge.
		const dr = (thickness / GUIDE_SCALE) * Math.min(w, h)
		const [iwd2, ihd2] = [wd2 - dr, hd2 - dr]

		const at = (rx: number, ry: number, angle: number) => {
			const point = onEllipse(rx, ry, angle)
			return { x: coord(wd2 + point.x), y: coord(hd2 + point.y) }
		}
		const start = at(wd2, hd2, stAng)
		const innerStart = at(iwd2, ihd2, istAng)

		return [
			`M ${start.x} ${start.y}`,
			arcTo(wd2, hd2, stAng, swAng, at),
			`L ${innerStart.x} ${innerStart.y}`,
			arcTo(iwd2, ihd2, istAng, -swAng, at),
			'Z',
		].join(' ')
	},
}

/** One `a:path`'s commands, scaled from its own `0..w`/`0..h` units into the box. */
function customPathData(path: GeometryPath, w: number, h: number): string {
	// `@w`/`@h` of 0 means "the path is already in the shape's coordinate space",
	// which is also the only sane reading when there is no denominator to divide by.
	const sx = path.w > 0 ? w / path.w : 1
	const sy = path.h > 0 ? h / path.h : 1
	const x = (value: number): number => value * sx
	const y = (value: number): number => value * sy

	// `arcTo` is relative to wherever the pen is, so the conversion needs to track
	// it: OOXML states an arc by its radii and its start/sweep angles, and SVG
	// states it by its radii and its *end point*.
	let penX = 0
	let penY = 0
	const parts: string[] = []

	for (const command of path.commands) {
		// The pen goes in as two numbers rather than as a thunk over the two `let`s
		// above. It was a closure only because `segmentOf` needed the value at call
		// time and `endOf` recomputed it; now the end point is computed once and
		// both take it.
		const moved = endOf(command, x, y, penX, penY)
		parts.push(segmentOf(command, x, y, moved))
		penX = moved.x
		penY = moved.y
	}
	return parts.join(' ')
}

type Scale = (value: number) => number

function segmentOf(command: GeometryCommand, x: Scale, y: Scale, end: { x: number; y: number }): string {
	switch (command.cmd) {
		case 'moveTo':
			return `M ${x(command.x)} ${y(command.y)}`
		case 'lnTo':
			return `L ${x(command.x)} ${y(command.y)}`
		case 'cubicBezTo':
			return `C ${x(command.x1)} ${y(command.y1)} ${x(command.x2)} ${y(command.y2)} ${x(command.x)} ${y(command.y)}`
		case 'quadBezTo':
			return `Q ${x(command.x1)} ${y(command.y1)} ${x(command.x)} ${y(command.y)}`
		case 'arcTo': {
			const largeArc = Math.abs(command.swAng) > 180 ? 1 : 0
			// Both systems measure a positive angle clockwise (y grows downward), so
			// a positive sweep is SVG's sweep-flag 1 with no sign juggling.
			const sweep = command.swAng >= 0 ? 1 : 0
			// The end point is `endOf`'s answer, not a second derivation of it. Two
			// spellings of the same formula with nothing checking they agree draws a
			// path whose next command starts where the previous one did not end.
			return `A ${x(command.wR)} ${y(command.hR)} 0 ${largeArc} ${sweep} ${end.x} ${end.y}`
		}
		case 'close':
			return 'Z'
	}
}

function endOf(
	command: GeometryCommand,
	x: Scale,
	y: Scale,
	penX: number,
	penY: number
): { x: number; y: number } {
	switch (command.cmd) {
		case 'moveTo':
		case 'lnTo':
		case 'cubicBezTo':
		case 'quadBezTo':
			return { x: x(command.x), y: y(command.y) }
		case 'arcTo': {
			// OOXML states an arc by its radii and its start/sweep angles; SVG states
			// it by its radii and its end point, so this is the conversion, and it is
			// the *only* place the end point is derived. The centre is wherever it has
			// to be for the pen to sit at `stAng` on it.
			const rx = x(command.wR)
			const ry = y(command.hR)
			const start = (command.stAng * Math.PI) / 180
			const finish = ((command.stAng + command.swAng) * Math.PI) / 180
			const cx = penX - rx * Math.cos(start)
			const cy = penY - ry * Math.sin(start)
			return { x: cx + rx * Math.cos(finish), y: cy + ry * Math.sin(finish) }
		}
		// A `close` returns the pen to the subpath's start. Tracking that exactly
		// would mean remembering every `moveTo`; the only command that reads the pen
		// is `arcTo`, and an `arcTo` straight after a `close` is not a shape any
		// source in scope produces.
		case 'close':
			return { x: penX, y: penY }
	}
}

/** Path data for a node's geometry, in local EMU space. */
export function pathOf(geometry: Geometry, w: number, h: number): PathResult {
	if (geometry.kind === 'custom') {
		const drawn = geometry.paths.filter((path) => path.commands.length > 0)
		if (drawn.length === 0) return { d: rect(w, h), fallback: 'custGeom' }
		return {
			d: drawn.map((path) => customPathData(path, w, h)).join(' '),
			// `@fill="none"` on every path means the freeform is an open stroke, and
			// filling it would paint over whatever it was drawn on top of.
			...(drawn.every((path) => path.fill === 'none') ? { unfilled: true } : {}),
			...(drawn.every((path) => !path.stroke) ? { unstroked: true } : {}),
		}
	}

	const preset = PRESETS[geometry.preset]
	if (preset === undefined) return { d: rect(w, h), fallback: geometry.preset }
	return { d: preset(w, h, geometry.adjustValues) }
}
