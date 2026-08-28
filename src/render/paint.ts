/**
 * {@link Fill} and {@link Stroke} → SVG paint.
 *
 * ## `inherit` paints nothing, on both arms
 *
 * `inherit` means **nothing this model can see states a paint at all**. It is
 * not a theme reference left unfollowed: the read model *does* resolve
 * `p:style/a:fillRef` and `a:lnRef`, through `resolvedFill` and `resolvedLine`,
 * and every shape they answer for has already arrived here as `solid`. What
 * reaches these two arms is the remainder — shapes carrying no style reference
 * for upstream to follow — so there is no value to look up and painting is a
 * decision rather than a lookup.
 *
 * Both arms make the same one: draw nothing. A shape that states no outline and
 * has no `a:lnRef` to fall through is one PowerPoint draws no outline for, and
 * the same holds of its interior.
 *
 * The fill arm painted a flat neutral until it was measured. The reason given was
 * that transparency would erase every themed shape — but themed shapes resolve to
 * `solid` and never reach this arm. Across 47 PowerPoint-authored decks every
 * shape that did reach it was a placeholder (title, body, slide number, footer,
 * date), none of which PowerPoint fills, so the neutral laid a grey rectangle
 * behind exactly the text a reader is meant to read, on nearly every slide with a
 * title.
 *
 * `none` is not certainly right either: a placeholder's interior can come from
 * the layout, which this model does not walk. It is the better guess, because an
 * absent fill distorts a slide far less than an invented one, and because it is
 * the only one of the two that adds nothing the deck did not say.
 *
 * The slide surface is the exception, and it is not decided here — a slide whose
 * whole chain states no background is white rather than transparent, which
 * `render/document.ts` handles where the surface is drawn.
 *
 * Neither choice touches the round trip: `Fill.inherit` and `Stroke.inherit`
 * travel through the island as themselves, and only the picture approximates.
 */

import { type Color, type DashStyle, EMU_PER_POINT, type Fill, type Gradient, type LineEnd, type Stroke } from '../ir/render'

/**
 * What a line is drawn in when it states a width or a dash but no colour — the
 * one thing `a:lnRef` would have supplied. Unlike an unstated fill this cannot be
 * resolved by drawing nothing: the deck says there is a line, so something has to
 * be visible, and the arm is marked approximate for it.
 *
 * Documented, not derived, and not observed: across 47 PowerPoint-authored decks
 * all 172 painted lines (117 shape outlines, 55 table borders) state their own
 * colour, so this is a defensive arm rather than a common one.
 */
const UNSTATED_LINE_COLOR = '#d8dce6'

// Re-exported so the renderer's own modules and its tests keep one import site
// for it; the definition is the model's, beside `EMU_PER_INCH`.
export { EMU_PER_POINT }

/**
 * Collects the `<defs>` a slide's paint needs, and hands out ids for them.
 *
 * Per slide rather than per document: ids only have to be unique within one
 * `<svg>`, and a counter that restarts each slide keeps them short and keeps a
 * one-slide change from renumbering every gradient in the deck.
 */
export class Defs {
	private readonly entries: string[] = []
	private next = 0

	add(build: (id: string) => string): string {
		const id = `pxh-p${this.next++}`
		this.entries.push(build(id))
		return `url(#${id})`
	}

	render(): string {
		return this.entries.length === 0 ? '' : `<defs>${this.entries.join('')}</defs>`
	}
}

/** A colour as CSS. The scheme arm paints its already-resolved hex. */
export function cssColor(color: Color): string {
	return color.kind === 'srgb' ? `#${color.hex}` : `#${color.effectiveHex}`
}

function gradientDef(gradient: Gradient, defs: Defs): string {
	const stops = gradient.stops
		.map((stop) => {
			const opacity = stop.color.alpha
			const extra = opacity === undefined ? '' : ` stop-opacity="${opacity}"`
			return `<stop offset="${stop.position}" stop-color="${cssColor(stop.color)}"${extra}/>`
		})
		.join('')

	if (gradient.kind === 'linear') {
		// Both systems measure the angle clockwise from the positive x-axis with y
		// growing downward, so the OOXML angle goes in unchanged. Rotating about the
		// bounding box's centre in `objectBoundingBox` units is what keeps the sweep
		// correct for a box that is not square.
		return defs.add(
			(id) =>
				`<linearGradient id="${id}" gradientTransform="rotate(${gradient.angleDeg} 0.5 0.5)">${stops}</linearGradient>`
		)
	}
	return defs.add((id) => `<radialGradient id="${id}">${stops}</radialGradient>`)
}

/**
 * A picture fill, as a pattern holding an `<image>` the hydration script fills
 * in. `preserveAspectRatio="none"` is the `stretch` semantics OOXML states.
 *
 * `tile` is drawn as a stretch and marked approximate: the IR records *that* the
 * fill tiles but not the tile's size, offset or alignment (`a:tile`'s
 * attributes), so a repeat drawn from what is here would be a repeat at an
 * invented scale — a wrong picture rather than a coarse one.
 */
function pictureDef(assetName: string, defs: Defs): string {
	return defs.add(
		(id) =>
			`<pattern id="${id}" width="1" height="1" patternContentUnits="objectBoundingBox">` +
			`<image data-pxh-asset="${escapeAttr(assetName)}" width="1" height="1" preserveAspectRatio="none"/>` +
			`</pattern>`
	)
}

export interface Painted {
	/** SVG presentation attributes, already serialized. */
	attrs: string
	/** Set when the paint is a documented guess rather than something the deck stated. */
	approx?: string
	/**
	 * False when this paint draws nothing: an explicit `a:noFill`, or nothing
	 * stated at all.
	 *
	 * A caller that needs to know this used to compare {@link attrs} against the
	 * string `'stroke="none"'`, which worked because two files agreed on a literal
	 * with no type relating them. Reorder the attributes, add a `stroke-opacity`,
	 * change the quotes, and every table border and picture outline silently starts
	 * drawing an invisible line over itself with nothing failing.
	 */
	painted: boolean
}

/** Attribute-value escaping. Shared by every attribute this module writes. */
export function escapeAttr(value: string): string {
	return value
		.replaceAll('&', '&amp;')
		.replaceAll('<', '&lt;')
		.replaceAll('>', '&gt;')
		.replaceAll('"', '&quot;')
}

/**
 * The two paints that draw nothing, named so a caller that has decided *not* to
 * paint states the same thing {@link fillPaint} and {@link strokePaint} state.
 * A preset geometry that declares itself unfilled or unstroked is such a caller.
 */
export const UNPAINTED_FILL: Painted = { attrs: 'fill="none"', painted: false }
export const UNPAINTED_STROKE: Painted = { attrs: 'stroke="none"', painted: false }

export function fillPaint(fill: Fill, defs: Defs): Painted {
	switch (fill.kind) {
		// Two different facts, one drawing. `none` is an explicit `a:noFill`; `inherit`
		// is the absence of any statement. They stay distinct in the IR and on the
		// island — only the paint coincides, because "nothing stated" and "nothing
		// wanted" look the same on a canvas.
		case 'none':
		case 'inherit':
			return { attrs: 'fill="none"', painted: false }
		case 'solid': {
			const opacity = fill.color.alpha
			return {
				attrs: `fill="${cssColor(fill.color)}"${opacity === undefined ? '' : ` fill-opacity="${opacity}"`}`,
				painted: true,
			}
		}
		case 'gradient':
			return { attrs: `fill="${gradientDef(fill.gradient, defs)}"`, painted: true }
		case 'picture': {
			const opacity = fill.alpha
			return {
				attrs: `fill="${pictureDef(fill.asset.$asset, defs)}"${opacity === undefined ? '' : ` fill-opacity="${opacity}"`}`,
				painted: true,
				...(fill.mode === 'tile' ? { approx: 'fill:tile' } : {}),
			}
		}
		case 'pattern': {
			// A preset hatch (`a:pattFill/@prst`) is 54 named bitmaps. Painting the
			// background colour keeps the shape's weight and mass right and leaves the
			// texture out, which is a coarse picture; inventing a hatch would be a
			// wrong one.
			const background = fill.background ?? fill.foreground
			return {
				attrs: background === null ? 'fill="none"' : `fill="${cssColor(background)}"`,
				painted: background !== null,
				approx: 'fill:pattern',
			}
		}
	}
}

export function strokePaint(stroke: Stroke, defs: Defs): Painted {
	if (stroke.kind === 'none' || stroke.kind === 'inherit') return { attrs: 'stroke="none"', painted: false }

	const parts: string[] = []
	const approx: string[] = []
	let markerPaint: string
	let markerOpacity: number | undefined
	if (stroke.gradient !== undefined) {
		markerPaint = gradientDef(stroke.gradient, defs)
		parts.push(`stroke="${markerPaint}"`)
	}
	else if (stroke.color !== undefined) {
		markerPaint = cssColor(stroke.color)
		markerOpacity = stroke.color.alpha
		parts.push(`stroke="${markerPaint}"`)
		if (markerOpacity !== undefined) parts.push(`stroke-opacity="${markerOpacity}"`)
	} else {
		markerPaint = UNSTATED_LINE_COLOR
		parts.push(`stroke="${markerPaint}"`)
		approx.push('line:color')
	}

	const widthEmu = (stroke.widthPt ?? 1) * EMU_PER_POINT
	parts.push(`stroke-width="${Math.round(widthEmu)}"`)

	if (stroke.cap !== undefined) parts.push(`stroke-linecap="${SVG_CAP[stroke.cap]}"`)
	if (stroke.dash !== undefined) parts.push(`stroke-dasharray="${dashArray(stroke.dash, widthEmu)}"`)

	const head = lineEndDef(stroke.head, markerPaint, markerOpacity, defs)
	if (head !== undefined) parts.push(`marker-start="${head}"`)
	const tail = lineEndDef(stroke.tail, markerPaint, markerOpacity, defs)
	if (tail !== undefined) parts.push(`marker-end="${tail}"`)

	return {
		attrs: parts.join(' '),
		painted: true,
		...(approx.length === 0 ? {} : { approx: approx.join(' ') }),
	}
}

/**
 * Small/medium/large line-end classes as factors of the medium marker.
 *
 * `LineEnd['width']` stands in for both axes: OOXML gives `@w` and `@len` the
 * same three classes, so one spelling covers a caller passing either.
 */
function lineEndScale(size: LineEnd['width']): number {
	if (size === 'sm') return 0.5
	if (size === 'lg') return 1.5
	return 1
}

/**
 * One DrawingML line end as an SVG marker.
 *
 * Marker units are stroke widths, matching OOXML's own definition of `@w` and
 * `@len` as sizes relative to the line. `auto-start-reverse` lets one rightward
 * definition serve both `a:headEnd` and `a:tailEnd`; SVG reverses the start arm
 * around the same tip without a second set of geometry.
 */
function lineEndDef(end: LineEnd | undefined, paint: string, opacity: number | undefined, defs: Defs): string | undefined {
	if (end === undefined || end.type === 'none') return undefined

	const markerWidth = 3 * lineEndScale(end.length)
	const markerHeight = 2.5 * lineEndScale(end.width)
	const alpha = opacity === undefined ? '' : ` opacity="${opacity}"`
	return defs.add((id) => {
		const marker =
			`<marker id="${id}" viewBox="0 0 10 10" refX="10" refY="5" ` +
			`markerWidth="${markerWidth}" markerHeight="${markerHeight}" markerUnits="strokeWidth" ` +
			'orient="auto-start-reverse" overflow="visible">'
		switch (end.type) {
			case 'triangle':
				return `${marker}<path d="M 0 0 L 10 5 L 0 10 Z" fill="${paint}"${alpha}/></marker>`
			case 'stealth':
				return `${marker}<path d="M 0 0 L 10 5 L 0 10 L 3 5 Z" fill="${paint}"${alpha}/></marker>`
			case 'diamond':
				return `${marker}<path d="M 0 5 L 5 0 L 10 5 L 5 10 Z" fill="${paint}"${alpha}/></marker>`
			case 'oval':
				return `${marker}<ellipse cx="5" cy="5" rx="5" ry="4" fill="${paint}"${alpha}/></marker>`
			case 'arrow':
				return `${marker}<path d="M 0 0 L 10 5 L 0 10" fill="none" stroke="${paint}" stroke-width="3"${alpha}/></marker>`
			case 'none':
				// Guarded before the def is allocated; retained for exhaustiveness across
				// the callback boundary, where TypeScript does not preserve that narrowing.
				return `${marker}</marker>`
		}
	})
}

/**
 * Three decimals, which is the rounding this whole channel uses and the one place
 * the reason for it is written.
 *
 * The unit varies by caller and the argument does not: everything here is either
 * EMU, a 914,400th of an inch, or a point. Three decimals of either is already far
 * below what a screen can show, and the alternative is output whose noise obscures
 * what it draws.
 */
export function round3(value: number): number {
	return Math.round(value * 1000) / 1000
}

const SVG_CAP: Record<'flat' | 'rnd' | 'sq', string> = { flat: 'butt', rnd: 'round', sq: 'square' }

/** The eleven `a:prstDash` presets, in multiples of the stroke width. */
const DASH_MULTIPLES: Record<DashStyle, number[]> = {
	solid: [],
	dot: [1, 3],
	dash: [4, 3],
	lgDash: [8, 3],
	dashDot: [4, 3, 1, 3],
	lgDashDot: [8, 3, 1, 3],
	lgDashDotDot: [8, 3, 1, 3, 1, 3],
	sysDash: [3, 1],
	sysDot: [1, 1],
	sysDashDot: [3, 1, 1, 1],
	sysDashDotDot: [3, 1, 1, 1, 1, 1],
}

/**
 * `a:prstDash` → `stroke-dasharray`, in multiples of the stroke width.
 *
 * That is how OOXML defines the preset dashes, and it is why the pattern has to
 * be computed from the width rather than written as constants: the same `dash`
 * on a 1pt and a 6pt line is a different array in absolute units.
 */
function dashArray(dash: DashStyle, widthEmu: number): string {
	const multiples = DASH_MULTIPLES[dash]
	if (multiples.length === 0) return 'none'
	return multiples.map((multiple) => Math.max(1, Math.round(multiple * widthEmu))).join(' ')
}
