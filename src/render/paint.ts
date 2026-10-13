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

import {
	type Color,
	type DashStyle,
	EMU_PER_POINT,
	type Fill,
	type Gradient,
	type LineEnd,
	type PictureEffect,
	type Shadow,
	type Stroke,
} from '../ir/render'

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

/** {@link cssColor} with the colour's own alpha, for the places CSS takes no separate opacity. */
function cssColorAlpha(color: Color): string {
	const hex = cssColor(color)
	if (color.alpha === undefined) return hex
	const channel = (at: number) => Number.parseInt(hex.slice(at, at + 2), 16)
	return `rgba(${channel(1)},${channel(3)},${channel(5)},${round3(color.alpha)})`
}

/**
 * A shadow's offset as x/y, in whatever unit `unit` converts a point to.
 *
 * The node is drawn inside its own rotation, so a shadow that rotates with the
 * shape needs nothing more; one that does not (`@rotWithShape="0"`) is turned back
 * by the shape's rotation to keep its direction on the slide.
 */
function shadowOffset(shadow: Shadow, rotation: number, unit: number): { dx: number; dy: number } {
	const turn = shadow.rotateWithShape === false ? rotation : 0
	const radians = (((shadow.angleDeg ?? 0) - turn) * Math.PI) / 180
	const distance = (shadow.offsetPt ?? 0) * unit
	return { dx: round3(distance * Math.cos(radians)), dy: round3(distance * Math.sin(radians)) }
}

/** An outer shadow as an SVG filter on a node drawn in EMU, returning its `url(#…)`. */
export function shadowFilter(shadow: Shadow, rotation: number, defs: Defs): string {
	const { dx, dy } = shadowOffset(shadow, rotation, EMU_PER_POINT)
	// `@blurRad` is a radius; a Gaussian's standard deviation is about half of one.
	const blur = round3(((shadow.blurPt ?? 0) * EMU_PER_POINT) / 2)
	const opacity = shadow.color.alpha === undefined ? '' : ` flood-opacity="${round3(shadow.color.alpha)}"`
	return defs.add(
		(id) =>
			`<filter id="${id}" x="-50%" y="-50%" width="200%" height="200%">` +
			`<feDropShadow dx="${dx}" dy="${dy}" stdDeviation="${blur}" flood-color="${cssColor(shadow.color)}"${opacity}/>` +
			`</filter>`
	)
}

/** An outer shadow on a run, as a CSS `text-shadow` in the frame's point-sized `px`. */
export function cssTextShadow(shadow: Shadow): string {
	const { dx, dy } = shadowOffset(shadow, 0, 1)
	return `${dx}px ${dy}px ${round3(shadow.blurPt ?? 0)}px ${cssColorAlpha(shadow.color)}`
}

/** A colour's three channels as 0–1 fractions, in sRGB. */
function channels(color: Color): [number, number, number] {
	const hex = cssColor(color)
	return [1, 3, 5].map((at) => Number.parseInt(hex.slice(at, at + 2), 16) / 255) as [number, number, number]
}

/** Luminance into all three channels, which is what every recolour below starts from. */
const GRAY = '0.299 0.587 0.114 0 0 0.299 0.587 0.114 0 0 0.299 0.587 0.114 0 0 0 0 0 1 0'

/**
 * One blip effect as filter primitives reading `input` and writing `output`.
 * Not `alpha`, which is an opacity and not a filter.
 */
function effectPrimitives(effect: Exclude<PictureEffect, { kind: 'alpha' }>, input: string, output: string): string {
	const io = `in="${input}" result="${output}"`
	switch (effect.kind) {
		case 'grayscale':
			return `<feColorMatrix ${io} type="matrix" values="${GRAY}"/>`
		case 'biLevel': {
			// A step at the threshold: a steep enough slope clamps every channel to 0 or 1.
			const slope = 1000
			const func = `type="linear" slope="${slope}" intercept="${round3(-slope * effect.threshold)}"`
			return (
				`<feColorMatrix in="${input}" result="${output}-gray" type="matrix" values="${GRAY}"/>` +
				`<feComponentTransfer in="${output}-gray" result="${output}">` +
				`<feFuncR ${func}/><feFuncG ${func}/><feFuncB ${func}/></feComponentTransfer>`
			)
		}
		case 'duotone': {
			const dark = channels(effect.dark)
			const light = channels(effect.light)
			const func = (i: number) => `type="table" tableValues="${round3(dark[i] as number)} ${round3(light[i] as number)}"`
			return (
				`<feColorMatrix in="${input}" result="${output}-gray" type="matrix" values="${GRAY}"/>` +
				`<feComponentTransfer in="${output}-gray" result="${output}">` +
				`<feFuncR ${func(0)}/><feFuncG ${func(1)}/><feFuncB ${func(2)}/></feComponentTransfer>`
			)
		}
		case 'clrChange': {
			// Each channel is 1 where it equals `from`'s byte and 0 elsewhere; an opaque
			// pixel whose three channels all hit is the one `a:clrChange` repaints.
			const from = channels(effect.from).map((value) => Math.round(value * 255))
			const table = (byte: number) =>
				`type="discrete" tableValues="${Array.from({ length: 256 }, (_, at) => (at === byte ? 1 : 0)).join(' ')}"`
			const opacity = effect.to.alpha === undefined ? '' : ` flood-opacity="${round3(effect.to.alpha)}"`
			return (
				`<feComponentTransfer in="${input}" result="${output}-hit">` +
				`<feFuncR ${table(from[0] as number)}/><feFuncG ${table(from[1] as number)}/><feFuncB ${table(from[2] as number)}/>` +
				`</feComponentTransfer>` +
				`<feColorMatrix in="${output}-hit" result="${output}-mask" type="matrix" values="0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 1 1 1 1 -3"/>` +
				`<feFlood result="${output}-to" flood-color="${cssColor(effect.to)}"${opacity}/>` +
				`<feComposite in="${output}-to" in2="${output}-mask" operator="in" result="${output}-swapped"/>` +
				`<feComposite in="${input}" in2="${output}-mask" operator="out" result="${output}-kept"/>` +
				`<feMerge result="${output}"><feMergeNode in="${output}-kept"/><feMergeNode in="${output}-swapped"/></feMerge>`
			)
		}
	}
}

/**
 * A picture's blip effects, as the attributes its `<image>` takes: one filter
 * chaining the recolours in document order, and an opacity for `a:alphaModFix`.
 * The filter works in sRGB, the space the effects' colours and thresholds are in.
 */
export function pictureEffects(effects: readonly PictureEffect[], defs: Defs): string {
	const attrs: string[] = []
	const primitives: string[] = []
	let input = 'SourceGraphic'
	for (const effect of effects) {
		if (effect.kind === 'alpha') {
			attrs.push(`opacity="${round3(effect.amount)}"`)
			continue
		}
		const output = `e${primitives.length}`
		primitives.push(effectPrimitives(effect, input, output))
		input = output
	}
	if (primitives.length > 0) {
		const filter = defs.add(
			(id) => `<filter id="${id}" color-interpolation-filters="sRGB">${primitives.join('')}</filter>`
		)
		attrs.push(`filter="${filter}"`)
	}
	return attrs.join(' ')
}

/**
 * What a gradient needs from the node it paints: the box's proportions, for an
 * angle that is not to be stretched with it, and the node's rotation, for a
 * gradient that does not turn with it.
 */
export interface PaintFrame {
	w: number
	h: number
	rotation: number
}

/**
 * A linear gradient's end points in `objectBoundingBox` units.
 *
 * DrawingML sweeps the gradient so the two farthest corners land on the first
 * and last stop. `a:lin/@scaled` says which space the angle is in: scaled, the
 * unit square, stretched with the box; unscaled, the box itself, so 45° is 45° on
 * the slide. The vector below is the one whose isolines in the stretched box are
 * perpendicular to the angle in that space, scaled so the corners reach the ends.
 */
function linearEnds(angleDeg: number, scaled: boolean, frame: PaintFrame): [number, number, number, number] {
	const radians = (angleDeg * Math.PI) / 180
	const dx = Math.cos(radians)
	const dy = Math.sin(radians)
	const [sx, sy] = scaled || frame.w <= 0 || frame.h <= 0 ? [1, 1] : [frame.w, frame.h]
	const [vx, vy] = [dx * sx, dy * sy]
	const k = (Math.abs(dx) * sx + Math.abs(dy) * sy) / (vx * vx + vy * vy)
	const [hx, hy] = [(vx * k) / 2, (vy * k) / 2]
	return [round3(0.5 - hx), round3(0.5 - hy), round3(0.5 + hx), round3(0.5 + hy)]
}

/**
 * A run's glow as `text-shadow` layers. OOXML's glow is the glyph grown by its
 * radius and softened at the edge; one blurred shadow of that radius fades out
 * long before it, so a second at half the radius gives the halo its body.
 */
export function cssTextGlow(glow: { color: Color; radiusPt?: number }): string[] {
	const radius = glow.radiusPt ?? 0
	if (radius <= 0) return []
	const color = cssColorAlpha(glow.color)
	return [`0 0 ${round3(radius / 2)}px ${color}`, `0 0 ${round3(radius)}px ${color}`]
}

function gradientDef(gradient: Gradient, frame: PaintFrame, defs: Defs): string {
	const stops = gradient.stops
		.map((stop) => {
			const opacity = stop.color.alpha
			const extra = opacity === undefined ? '' : ` stop-opacity="${opacity}"`
			return `<stop offset="${stop.position}" stop-color="${cssColor(stop.color)}"${extra}/>`
		})
		.join('')

	if (gradient.kind === 'linear') {
		// Both systems measure the angle clockwise from the positive x-axis with y
		// growing downward, so the OOXML angle goes in unchanged. A gradient that does
		// not rotate with its shape is turned back by the shape's rotation, since it
		// is drawn inside the shape's transform.
		const turn = gradient.rotateWithShape === false ? frame.rotation : 0
		const [x1, y1, x2, y2] = linearEnds(gradient.angleDeg - turn, gradient.scaled === true, frame)
		return defs.add(
			(id) => `<linearGradient id="${id}" x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}">${stops}</linearGradient>`
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

export function fillPaint(fill: Fill, frame: PaintFrame, defs: Defs): Painted {
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
			return { attrs: `fill="${gradientDef(fill.gradient, frame, defs)}"`, painted: true }
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

export function strokePaint(stroke: Stroke, frame: PaintFrame, defs: Defs): Painted {
	if (stroke.kind === 'none' || stroke.kind === 'inherit') return { attrs: 'stroke="none"', painted: false }

	const parts: string[] = []
	const approx: string[] = []
	let markerPaint: string
	let markerOpacity: number | undefined
	if (stroke.gradient !== undefined) {
		markerPaint = gradientDef(stroke.gradient, frame, defs)
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
