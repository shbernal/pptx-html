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
 * So this is a **picture**, and it is allowed to be an incomplete one. The IR
 * carries `preset` and `adjustValues` verbatim through the island, so a shape
 * this file cannot draw still round-trips byte-for-byte — it just looks like a
 * rectangle in the preview. That is the difference between the two channels
 * doing their jobs and a silent loss, and it is why the unresolved case is
 * *marked* ({@link PathResult.fallback}) rather than quietly boxed.
 */

import type { Geometry, GeometryCommand, GeometryPath } from '../ir/render'

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
 */
const PRESETS: Record<string, (w: number, h: number, adj: Record<string, string>) => string> = {
	rect,
	// A picture's default geometry, and the shape every unresolved preset becomes.
	roundRect: (w, h, adj) => {
		const radius = (adjust(adj, 'adj', 16_667) / GUIDE_SCALE) * Math.min(w, h)
		const r = Math.max(0, Math.min(radius, Math.min(w, h) / 2))
		return [
			`M ${r} 0`,
			`L ${w - r} 0`,
			`A ${r} ${r} 0 0 1 ${w} ${r}`,
			`L ${w} ${h - r}`,
			`A ${r} ${r} 0 0 1 ${w - r} ${h}`,
			`L ${r} ${h}`,
			`A ${r} ${r} 0 0 1 0 ${h - r}`,
			`L 0 ${r}`,
			`A ${r} ${r} 0 0 1 ${r} 0`,
			'Z',
		].join(' ')
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
		parts.push(segmentOf(command, x, y, () => ({ penX, penY })))
		const moved = endOf(command, x, y, penX, penY)
		penX = moved.x
		penY = moved.y
	}
	return parts.join(' ')
}

type Scale = (value: number) => number

function segmentOf(command: GeometryCommand, x: Scale, y: Scale, pen: () => { penX: number; penY: number }): string {
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
			const { penX, penY } = pen()
			const rx = x(command.wR)
			const ry = y(command.hR)
			const start = (command.stAng * Math.PI) / 180
			const end = ((command.stAng + command.swAng) * Math.PI) / 180
			// The centre is wherever it has to be for the pen to sit at `stAng` on it.
			const cx = penX - rx * Math.cos(start)
			const cy = penY - ry * Math.sin(start)
			const largeArc = Math.abs(command.swAng) > 180 ? 1 : 0
			// Both systems measure a positive angle clockwise (y grows downward), so
			// a positive sweep is SVG's sweep-flag 1 with no sign juggling.
			const sweep = command.swAng >= 0 ? 1 : 0
			return `A ${rx} ${ry} 0 ${largeArc} ${sweep} ${cx + rx * Math.cos(end)} ${cy + ry * Math.sin(end)}`
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
			const rx = x(command.wR)
			const ry = y(command.hR)
			const start = (command.stAng * Math.PI) / 180
			const end = ((command.stAng + command.swAng) * Math.PI) / 180
			return {
				x: penX - rx * Math.cos(start) + rx * Math.cos(end),
				y: penY - ry * Math.sin(start) + ry * Math.sin(end),
			}
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
