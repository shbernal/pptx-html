/**
 * Read model → IR paint: colour, gradient, fill, stroke.
 *
 * One file rather than the `color.ts` the outline called for, because a `Fill` is
 * made of `Color`s and a `Stroke` is made of both, and splitting them puts three
 * halves of one decode in three places. The rules they share are the reason:
 *
 * - **A colour keeps all three of its parts.** The scheme token is what the deck
 *   said, the transform list is how it was modified, and `effectiveHex` is what a
 *   renderer paints. Keeping only the last is the flattening that makes a theme
 *   change stop propagating; keeping only the first two leaves nothing to draw.
 * - **Absent is not `none`.** A shape that states no fill inherits one from its
 *   style reference or its placeholder; a shape with `a:noFill` is deliberately
 *   transparent. See {@link Fill}.
 * - **A value outside the modeled subset is a note, never a rounding.** An
 *   unmodeled dash token becomes a {@link FidelityNote} and an absent dash, not a
 *   silent `solid`.
 */

import type {
	CellBorder,
	GradientFill,
	GradientStop as ReadGradientStop,
	PatternFill,
	PictureFill,
	ResolvedColor,
} from '@shbernal/ts-pptx/read'
import type {
	Color,
	DashStyle,
	EdgeRect,
	Fill,
	Gradient,
	GradientStop,
	LineEnd,
	SchemeToken,
	Stroke,
} from '../ir/render'
import { type ImportScope, note } from './context'

/** The write API expresses these; every other `ST_PresetLineDashVal` is a note. */
const DASH_STYLES = new Set<string>([
	'solid',
	'dot',
	'dash',
	'lgDash',
	'dashDot',
	'lgDashDot',
	'lgDashDotDot',
	'sysDash',
	'sysDot',
	'sysDashDot',
	'sysDashDotDot',
])

const LINE_END_TYPES = new Set<string>(['none', 'triangle', 'stealth', 'diamond', 'oval', 'arrow'])
const LINE_END_SIZES = new Set<string>(['sm', 'med', 'lg'])

/**
 * A colour reference, from the read model's own split: the raw `a:schemeClr/@val`
 * token when there was one, and the resolution of whichever element was there.
 *
 * Both arguments, because neither alone is the colour. `resolved` carries the
 * transforms and the painted hex but not which token produced them; the token
 * carries the reference but nothing to draw.
 */
export function colorOf(token: string | null, resolved: ResolvedColor | null): Color | undefined {
	if (token !== null) {
		const color: Color = {
			kind: 'scheme',
			slot: token as SchemeToken,
			transforms: resolved?.transforms ?? [],
			effectiveHex: resolved?.effectiveHex ?? resolved?.hex ?? '000000',
		}
		return resolved?.alpha === undefined ? color : { ...color, alpha: resolved.alpha }
	}
	if (resolved === null) return undefined
	const color: Color = { kind: 'srgb', hex: resolved.effectiveHex }
	return resolved.alpha === undefined ? color : { ...color, alpha: resolved.alpha }
}

/** A gradient stop, whose colour arrives pre-split by the read model. */
function stopOf(stop: ReadGradientStop): GradientStop | null {
	const hex = stop.effectiveHex ?? stop.color
	if (hex === null) return null
	const color: Color =
		stop.schemeColor === null
			? { kind: 'srgb', hex }
			: { kind: 'scheme', slot: stop.schemeColor as SchemeToken, transforms: [], effectiveHex: hex }
	return {
		position: stop.position ?? 0,
		color: stop.alpha === undefined ? color : ({ ...color, alpha: stop.alpha } as Color),
	}
}

export function gradientOf(fill: GradientFill, scope: ImportScope): Gradient | null {
	const stops = fill.stops.map(stopOf).filter((stop): stop is GradientStop => stop !== null)
	if (stops.length === 0) return null
	if (fill.kind === 'path') {
		const shape = fill.path === 'circle' || fill.path === 'rect' || fill.path === 'shape' ? fill.path : null
		if (shape === null) {
			note(
				scope,
				'fill.gradient.path',
				'approximated',
				'unsupported',
				`a path gradient whose shape is ${JSON.stringify(fill.path)} is drawn as a circular one`
			)
			return { kind: 'path', shape: 'circle', stops }
		}
		return { kind: 'path', shape, stops }
	}
	return { kind: 'linear', angleDeg: fill.angleDeg ?? 0, stops }
}

function edgeRectOf(rect: { left: number; top: number; right: number; bottom: number }): EdgeRect {
	return { left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom }
}

/**
 * The fill-bearing members of a shape, a table or a table cell.
 *
 * Structural rather than a union of the three classes: they already agree on
 * these names, and naming the classes would make this file import the whole read
 * model to describe five getters.
 */
export interface FillSource {
	readonly fillSchemeColor: string | null
	readonly resolvedFill: ResolvedColor | null
	readonly gradientFill: GradientFill | null
	readonly patternFill: PatternFill | null
	readonly pictureFill: PictureFill | null
	/**
	 * Optional because `Shape` has it and `TableCell` does not — the one member of
	 * this set the two classes disagree about. Optional rather than two interfaces:
	 * an absent reader and a `false` one mean different things here, and a cell
	 * silently satisfying `fillNoFill: boolean` by structural luck is worse than
	 * the arm below being unreachable for it.
	 */
	readonly fillNoFill?: boolean
}

export function fillOf(source: FillSource, scope: ImportScope): Fill {
	// First because `EG_FillProperties` admits one child: a shape with `a:noFill`
	// has no other fill for the branches below to find.
	if (source.fillNoFill === true) return { kind: 'none' }

	const picture = source.pictureFill
	if (picture !== null) return pictureFillOf(picture, scope)

	const gradient = source.gradientFill
	if (gradient !== null) {
		const mapped = gradientOf(gradient, scope)
		if (mapped !== null) return { kind: 'gradient', gradient: mapped }
	}

	const pattern = source.patternFill
	if (pattern !== null) {
		return {
			kind: 'pattern',
			preset: pattern.preset,
			foreground: colorOf(null, pattern.foreground) ?? null,
			background: colorOf(null, pattern.background) ?? null,
		}
	}

	const solid = colorOf(source.fillSchemeColor, source.resolvedFill)
	if (solid !== undefined) return { kind: 'solid', color: solid }

	// A shape has already taken the `none` arm above. A table cell cannot: `TableCell`
	// exposes `hasOwnFill` but no `noFill` reader, so an explicit `a:noFill` on a cell
	// is indistinguishable from one inheriting the table style's shading, and both
	// land here as inherited — the truthful half of the two.
	//
	// Deriving it from `hasOwnFill && every other accessor is null` would be a guess,
	// not a read: the same shape is produced by any fill choice this decode does not
	// model. The ask is the cell-side counterpart of `Shape.fillNoFill`:
	// https://github.com/shbernal/ts-pptx/issues/7
	return { kind: 'inherit' }
}

export function pictureFillOf(picture: PictureFill, scope: ImportScope): Fill {
	const asset = scope.assets.refFor(picture.partName)
	if (asset === null) {
		note(
			scope,
			'fill.picture',
			'dropped',
			'unread',
			'this picture fill names no resolvable media part, so there are no bytes to paint with'
		)
		return { kind: 'inherit' }
	}
	const fill: Fill = { kind: 'picture', asset, mode: picture.mode === 'tile' ? 'tile' : 'stretch' }
	const withRect = picture.srcRect === null ? fill : { ...fill, srcRect: edgeRectOf(picture.srcRect) }
	return picture.alpha === null ? withRect : { ...withRect, alpha: picture.alpha }
}

/** The line-bearing members of a shape. Structural, for the reason {@link FillSource} is. */
export interface StrokeSource {
	readonly lineNoFill: boolean
	readonly lineWidthPt: number | null
	readonly lineDash: string | null
	readonly lineCap: string | null
	readonly lineAlign: string | null
	readonly lineSchemeColor: string | null
	readonly resolvedLine: ResolvedColor | null
	readonly lineGradient: GradientFill | null
	readonly lineEnds: { head: unknown; tail: unknown } | null
}

interface ReadLineEnd {
	type: string
	width: string | null
	length: string | null
}

function lineEndOf(end: unknown, scope: ImportScope): LineEnd | undefined {
	if (end === null || end === undefined) return undefined
	const source = end as ReadLineEnd
	if (!LINE_END_TYPES.has(source.type)) {
		note(
			scope,
			'line.arrowSize',
			'dropped',
			'unsupported',
			`an arrowhead of type ${JSON.stringify(source.type)} is outside the modeled set and is not drawn`
		)
		return undefined
	}
	const result: LineEnd = { type: source.type as LineEnd['type'] }
	const width = source.width !== null && LINE_END_SIZES.has(source.width) ? (source.width as 'sm' | 'med' | 'lg') : null
	const length =
		source.length !== null && LINE_END_SIZES.has(source.length) ? (source.length as 'sm' | 'med' | 'lg') : null
	return {
		...result,
		...(width === null ? {} : { width }),
		...(length === null ? {} : { length }),
	}
}

/**
 * `construct` is a parameter because upstream keys a cell edge's dash separately
 * from a shape line's (`table.cell.borders.dash` vs `line.dash`), and a note
 * filed under the wrong one of those two is a note that can never match.
 */
function dashOf(dash: string | null, scope: ImportScope, construct: string): DashStyle | undefined {
	if (dash === null) return undefined
	if (DASH_STYLES.has(dash)) return dash as DashStyle
	note(
		scope,
		construct,
		'dropped',
		'unsupported',
		`a dash style of ${JSON.stringify(dash)} has no modeled equivalent, so the line is drawn without one`
	)
	return undefined
}

/**
 * `a:ln/@cap`, which the write API can carry back (`ShapeLineProps.cap`), and
 * `@algn`, which it cannot.
 *
 * Both are read as of ts-pptx 3.0.0. The split between modeling one and noting
 * the other is the round-trip rule, not a reading limit: a field that could never
 * come back would be a difference the model quietly absorbs instead of declaring.
 *
 * `line.cap` and `line.align` are coined keys — upstream's set has `line.dash`,
 * `line.width` and `line.arrowSize` and reaches neither of these. Coining is for
 * a construct upstream does not model at all, which is exactly the case here;
 * the rule it does not break is inventing a *synonym* for a key that exists.
 */
const LINE_CAPS = new Set<string>(['flat', 'rnd', 'sq'])

function capOf(cap: string | null, scope: ImportScope): 'flat' | 'rnd' | 'sq' | undefined {
	if (cap === null) return undefined
	if (LINE_CAPS.has(cap)) return cap as 'flat' | 'rnd' | 'sq'
	note(
		scope,
		'line.cap',
		'dropped',
		'unsupported',
		`a line cap of ${JSON.stringify(cap)} is outside ST_LineCap, so the line is drawn with the default flat cap`
	)
	return undefined
}

function noteLineAlign(align: string | null, scope: ImportScope): void {
	if (align === null) return
	note(
		scope,
		'line.align',
		'dropped',
		'unwritable',
		`this outline states a:ln/@algn of ${JSON.stringify(align)}; the write API has no option for it, so the outline comes back centred on the shape's edge and a thick one sits half its width further out`
	)
}

export function strokeOf(source: StrokeSource, scope: ImportScope): Stroke {
	if (source.lineNoFill) return { kind: 'none' }

	const color = colorOf(source.lineSchemeColor, source.resolvedLine)
	const gradient = source.lineGradient === null ? null : gradientOf(source.lineGradient, scope)
	const dash = dashOf(source.lineDash, scope, 'line.dash')
	const cap = capOf(source.lineCap, scope)
	const head = lineEndOf(source.lineEnds?.head, scope)
	const tail = lineEndOf(source.lineEnds?.tail, scope)
	noteLineAlign(source.lineAlign, scope)

	// Nothing modeled stated at all: the line comes from the theme's `a:lnRef`, which
	// the read model does not resolve. Reporting `none` here would erase every themed
	// outline in the deck. `@algn` is deliberately not part of this test — it produced
	// a note and nothing else, so a line stating only `@algn` still states nothing this
	// model carries.
	if (
		color === undefined &&
		gradient === null &&
		dash === undefined &&
		cap === undefined &&
		head === undefined &&
		tail === undefined &&
		source.lineWidthPt === null
	) {
		return { kind: 'inherit' }
	}

	return {
		kind: 'line',
		...(source.lineWidthPt === null ? {} : { widthPt: source.lineWidthPt }),
		...(color === undefined ? {} : { color }),
		...(gradient === null ? {} : { gradient }),
		...(dash === undefined ? {} : { dash }),
		...(cap === undefined ? {} : { cap }),
		...(head === undefined ? {} : { head }),
		...(tail === undefined ? {} : { tail }),
	}
}

/**
 * A table cell edge. Its own decode because `CellBorder` is a different shape
 * from a shape's line accessors — the colour arrives already resolved to a hex
 * rather than as a {@link ResolvedColor}, so there are no transforms to keep.
 */
export function cellBorderOf(border: CellBorder | null, scope: ImportScope): Stroke {
	if (border === null) return { kind: 'inherit' }
	if (border.noFill) return { kind: 'none' }

	const color: Color | undefined =
		border.schemeColor !== null
			? {
					kind: 'scheme',
					slot: border.schemeColor as SchemeToken,
					transforms: [],
					effectiveHex: border.color ?? '000000',
				}
			: border.color !== null
				? { kind: 'srgb', hex: border.color }
				: undefined
	const dash = dashOf(border.dash, scope, 'table.cell.borders.dash')

	if (color === undefined && dash === undefined && border.widthPt === null) return { kind: 'inherit' }
	return {
		kind: 'line',
		...(border.widthPt === null ? {} : { widthPt: border.widthPt }),
		...(color === undefined ? {} : { color }),
		...(dash === undefined ? {} : { dash }),
	}
}
