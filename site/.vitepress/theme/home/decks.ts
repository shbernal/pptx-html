/**
 * The two decks the landing page puts in motion.
 *
 * These are **written by `@shbernal/ts-pptx` at page load**, imported, and drawn
 * by `renderDeck` in the visitor's browser. Nothing here is a picture of a slide,
 * and there is no image of a deck anywhere on this site — a screenshot would
 * prove exactly the thing the project refuses to do (see docs/decisions). What
 * moves across the home page is the renderer's own output or it is nothing.
 *
 * ## Why these are not corpus decks
 *
 * `test/corpus/decks.ts` is the round-trip oracle's input domain: one deck per
 * construct, deliberately minimal, and every one of them gated. The playground
 * shows those, because the playground makes a claim about fidelity. This page
 * makes a different and smaller claim — *this is what the renderer draws* — so
 * putting eight-slide consulting decks in the corpus to decorate a landing page
 * would widen the gate's scope for a presentational reason. They live here, and
 * the page says what they are.
 *
 * They are still inside the guarantee's stated input domain, because that domain
 * is "decks written by `@shbernal/ts-pptx`" and these are written by it.
 *
 * ## Drawn with what the renderer actually draws
 *
 * `src/render/geometry.ts` resolves nine presets exactly and boxes the rest —
 * marked as a fallback, on purpose, because a wrong outline is worse than an
 * obvious one. So every graphic below is built from those nine (`rect`,
 * `roundRect`, `ellipse`, `triangle`, `diamond`, `parallelogram`, `line`) or from
 * freeform `custGeom` paths, where the geometry is stated rather than named. The
 * charts are drawn as shapes for the same reason: `addChart` is a `hard`-tier
 * construct that takes the carried lane, and a carried chart draws as a
 * placeholder rather than a chart.
 *
 * Nothing here may depend on the clock, the network or randomness: the decks are
 * rebuilt on every page load and two visitors must get the same bytes.
 *
 * The companies, numbers and dates are fictitious, and the page says so.
 */

import { ShapeType, TsPptx } from '@shbernal/ts-pptx'

/** `LAYOUT_16x9`, in inches. Every coordinate below is stated against these. */
const W = 10
const H = 5.625

type Slide = ReturnType<TsPptx['addSlide']>

/** A monospaced-name palette, so a slide reads as a composition rather than hex. */
interface Palette {
	/** Darkest ground, for the full-bleed slides. */
	ink: string
	/** The second stop of every dark gradient. */
	deep: string
	accent: string
	accent2: string
	accent3: string
	paper: string
	/** The pale plate colour panels sit on. */
	mist: string
	/** Secondary type on a light ground. */
	muted: string
	/** Hairlines and table borders. */
	rule: string
	/** Type on a dark ground. */
	onDark: string
	/** Secondary type on a dark ground. */
	onDarkMuted: string
}

const INDIGO: Palette = {
	ink: '0C1030',
	deep: '221C6B',
	accent: '3D5AFE',
	accent2: '00C2CB',
	accent3: '8A7CFF',
	paper: 'FFFFFF',
	mist: 'EEF1FB',
	muted: '6B759E',
	rule: 'D9DFF2',
	onDark: 'FFFFFF',
	onDarkMuted: 'A8B0D8',
}

const PLUM: Palette = {
	ink: '1B0F26',
	deep: '54194A',
	accent: 'D81E5B',
	accent2: 'F5842F',
	accent3: 'FFC94A',
	paper: 'FFFFFF',
	mist: 'FBF1EA',
	muted: '7C6A86',
	rule: 'EBDCE4',
	onDark: 'FFFFFF',
	onDarkMuted: 'D5B9CC',
}

/**
 * A fill that paints nothing — an outline whose interior lets the slide through.
 *
 * `{ type: 'none' }` emits `<a:noFill/>` as of ts-pptx 3.1.0
 * ({@link https://github.com/shbernal/ts-pptx/issues/9}); before that it emitted
 * no fill child at all, which is the *inherit* state, and these shapes came back
 * painted in the theme's accent. The stopgap was an `a:solidFill` at zero alpha,
 * which paints the same but says something else — a white plate nobody can see
 * rather than no plate.
 *
 * Not merely omitting `fill`: on this writer omission still means `<a:noFill/>`
 * too, but by way of a ternary's else-arm rather than a stated intent
 * ({@link https://github.com/shbernal/ts-pptx/issues/10}).
 */
const CLEAR = { type: 'none' } as const

// ---------------------------------------------------------------------------
// Freeform geometry
// ---------------------------------------------------------------------------

/**
 * A `GeometryPoint` list, in the shape's own inch space.
 *
 * Typed structurally rather than imported: the writer's `GeometryPoint` is not
 * part of its public entry point, and the three node forms used here are small
 * enough to state.
 */
type Point =
	| { x: number; y: number; moveTo?: boolean }
	| { x: number; y: number; curve: { type: 'cubic'; x1: number; y1: number; x2: number; y2: number } }
	| { close: true }

const rad = (degrees: number): number => (degrees * Math.PI) / 180

/** A point on the ellipse. 0° is twelve o'clock and angles run clockwise. */
function onArc(cx: number, cy: number, rx: number, ry: number, degrees: number): { x: number; y: number } {
	return { x: cx + rx * Math.sin(rad(degrees)), y: cy - ry * Math.cos(rad(degrees)) }
}

/** d/dθ of {@link onArc}, which is what the Bézier handles are scaled from. */
function arcTangent(rx: number, ry: number, degrees: number): { x: number; y: number } {
	return { x: rx * Math.cos(rad(degrees)), y: ry * Math.sin(rad(degrees)) }
}

/**
 * An elliptical arc as cubic segments.
 *
 * The writer has an `arcTo` node and the renderer draws it, but stating the arc
 * as cubics keeps this file to the two path verbs whose output is identical on
 * both sides of the loop, and the error of the standard 4/3·tan(Δ/4) handle over
 * a quarter turn is under a thousandth of the radius — invisible at any size a
 * slide is drawn at.
 */
function arcPoints(cx: number, cy: number, rx: number, ry: number, from: number, to: number, move: boolean): Point[] {
	const steps = Math.max(1, Math.ceil(Math.abs(to - from) / 90))
	const step = (to - from) / steps
	const alpha = (4 / 3) * Math.tan(rad(step) / 4)
	const points: Point[] = []
	if (move) points.push({ ...onArc(cx, cy, rx, ry, from), moveTo: true })
	for (let index = 0; index < steps; index += 1) {
		const a = from + step * index
		const b = a + step
		const start = onArc(cx, cy, rx, ry, a)
		const end = onArc(cx, cy, rx, ry, b)
		const dStart = arcTangent(rx, ry, a)
		const dEnd = arcTangent(rx, ry, b)
		points.push({
			x: end.x,
			y: end.y,
			curve: {
				type: 'cubic',
				x1: start.x + alpha * dStart.x,
				y1: start.y + alpha * dStart.y,
				x2: end.x - alpha * dEnd.x,
				y2: end.y - alpha * dEnd.y,
			},
		})
	}
	return points
}

/** One ring segment of a donut, drawn in the box the whole donut occupies. */
function donutSlice(cx: number, cy: number, outer: number, inner: number, from: number, to: number): Point[] {
	return [
		...arcPoints(cx, cy, outer, outer, from, to, true),
		{ ...onArc(cx, cy, inner, inner, to) },
		...arcPoints(cx, cy, inner, inner, to, from, false),
		{ close: true },
	]
}

/** A polyline through `values`, closed down to `baseline` — the area under a trend. */
function areaPoints(points: readonly { x: number; y: number }[], baseline: number): Point[] {
	const first = points[0]
	const last = points[points.length - 1]
	if (first === undefined || last === undefined) throw new Error('an area needs at least one point')
	return [
		{ x: first.x, y: baseline, moveTo: true },
		...points.map((point) => ({ x: point.x, y: point.y })),
		{ x: last.x, y: baseline },
		{ close: true },
	]
}

// ---------------------------------------------------------------------------
// Slide furniture
// ---------------------------------------------------------------------------

/**
 * A slide, and a namer for it.
 *
 * `objectName` has to be unique within a slide — the writer reports
 * `object-name/duplicate` otherwise — and every helper below adds shapes without
 * knowing what else is on the page, so the counter is handed round rather than
 * left to each call site to get right.
 */
interface Page {
	slide: Slide
	/** A unique `objectName` under `base`. */
	id: (base: string) => string
}

function page(pptx: TsPptx): Page {
	const slide = pptx.addSlide()
	let seq = 0
	return {
		slide,
		id: (base) => {
			seq += 1
			return `${base}-${seq}`
		},
	}
}

/** The whole slide, in one flat colour. */
function ground(p: Page, color: string): void {
	p.slide.addShape('rect', { x: 0, y: 0, w: W, h: H, fill: { color }, objectName: p.id('ground') })
}

/** The whole slide, as a two-stop diagonal — the dark title and closing pages. */
function gradientGround(p: Page, from: string, to: string, angle = 35): void {
	p.slide.addShape('rect', {
		x: 0,
		y: 0,
		w: W,
		h: H,
		fill: {
			type: 'gradient',
			gradient: {
				kind: 'linear',
				angle,
				stops: [
					{ position: 0, color: from },
					{ position: 100, color: to },
				],
			},
		},
		objectName: p.id('ground'),
	})
}

interface TextOptions {
	size: number
	color: string
	bold?: boolean
	italic?: boolean
	align?: 'left' | 'center' | 'right'
	valign?: 'top' | 'middle' | 'bottom'
	face?: string
	spacing?: number
}

const BODY_FACE = 'Arial'

/**
 * A text box.
 *
 * Two things the caller does not have to remember, because getting either wrong
 * clips a line of type and the clip is only visible once the deck is drawn:
 *
 * - **`margin: 0`.** A text body's default insets are 0.05 in top and bottom, and
 *   the renderer applies them as padding inside a frame that is `overflow:
 *   hidden`. A 10.5 pt line in a 0.24 in box then loses its bottom half. With the
 *   insets off, `y` is where the type starts and `h` is the box.
 * - **A floor under `h`.** One line at `size` needs `size × 1.4` points of frame
 *   at the browser's default line height, so the height asked for is raised to
 *   that when it is smaller. Boxes that hold more than a line already ask for
 *   more and are untouched.
 */
function label(p: Page, text: string, x: number, y: number, w: number, h: number, options: TextOptions): void {
	p.slide.addText(text, {
		x,
		y,
		w,
		h: Math.max(h, (options.size / 72) * 1.4),
		margin: 0,
		fontSize: options.size,
		fontFace: options.face ?? BODY_FACE,
		color: options.color,
		bold: options.bold ?? false,
		italic: options.italic ?? false,
		align: options.align ?? 'left',
		valign: options.valign ?? 'top',
		charSpacing: options.spacing,
		objectName: p.id('copy'),
	})
}

/** The small tracked-out line above a headline. Every content slide has one. */
function eyebrow(p: Page, text: string, color: string, accent: string): void {
	p.slide.addShape('rect', { x: 0.62, y: 0.52, w: 0.34, h: 0.055, fill: { color: accent }, objectName: p.id('tick') })
	label(p, text.toUpperCase(), 0.62, 0.68, 6, 0.24, { size: 10.5, color, bold: true, spacing: 2.4 })
}

function headline(p: Page, text: string, color: string, size = 30): void {
	label(p, text, 0.6, 1.0, 8.8, 0.8, { size, color, bold: true })
}

function hairline(p: Page, x: number, y: number, w: number, color: string, weight = 0.75): void {
	p.slide.addShape('line', { x, y, w, h: 0, line: { color, width: weight }, objectName: p.id('rule') })
}

/** Slide number and running footer, on the light pages. */
function footer(p: Page, deck: string, number: number, palette: Palette): void {
	hairline(p, 0.6, H - 0.55, W - 1.2, palette.rule)
	label(p, deck.toUpperCase(), 0.6, H - 0.46, 6, 0.24, { size: 9, color: palette.muted, spacing: 1.2 })
	label(p, String(number), W - 1.2, H - 0.46, 0.6, 0.24, { size: 9, color: palette.muted, align: 'right' })
}

/** A rounded plate — the card every panel on these decks sits on. */
function plate(p: Page, x: number, y: number, w: number, h: number, color: string, radius = 0.06): void {
	p.slide.addShape('roundRect', {
		x,
		y,
		w,
		h,
		fill: { color },
		rectRadius: radius,
		objectName: p.id('plate'),
	})
}

// ---------------------------------------------------------------------------
// Deck one — Northwind Retail
// ---------------------------------------------------------------------------

const NORTHWIND = 'Northwind Retail · Growth review'

function northwindTitle(pptx: TsPptx, c: Palette): void {
	const p = page(pptx)
	gradientGround(p, c.ink, c.deep)

	// A freeform sweep off the right edge. Two arcs and a straight back, so the
	// curve is stated rather than approximated by a preset that would box.
	p.slide.addShape(ShapeType.custGeom, {
		x: 5.1,
		y: -1.2,
		w: 6.4,
		h: 8,
		fill: { color: c.accent, transparency: 62 },
		points: [
			{ x: 1.2, y: 0, moveTo: true },
			...arcPoints(1.2, 4, 3.4, 4, 0, 180, false),
			{ x: 3.4, y: 8 },
			...arcPoints(3.4, 4, 3.0, 4, 180, 0, false),
			{ close: true },
		],
		objectName: 'sweep',
	})
	p.slide.addShape(ShapeType.custGeom, {
		x: 6.4,
		y: -0.6,
		w: 5.2,
		h: 6.8,
		fill: { color: c.accent2, transparency: 58 },
		points: [
			{ x: 0.8, y: 0, moveTo: true },
			...arcPoints(0.8, 3.4, 2.6, 3.4, 0, 180, false),
			{ x: 2.4, y: 6.8 },
			...arcPoints(2.4, 3.4, 2.2, 3.4, 180, 0, false),
			{ close: true },
		],
		objectName: 'sweep-inner',
	})

	p.slide.addShape('rect', { x: 0.7, y: 1.5, w: 1.1, h: 0.07, fill: { color: c.accent2 }, objectName: 'kicker' })
	label(p, 'CONFIDENTIAL — BOARD REVIEW', 0.7, 0.72, 6, 0.26, {
		size: 10.5,
		color: c.onDarkMuted,
		bold: true,
		spacing: 2.6,
	})
	label(p, 'Northwind Retail', 0.66, 1.78, 6.4, 0.95, { size: 46, color: c.onDark, bold: true })
	label(p, 'Where the next £400m of growth comes from', 0.7, 2.78, 5.4, 0.9, { size: 17, color: c.onDarkMuted })
	hairline(p, 0.7, 3.95, 3.4, c.accent3, 1)
	label(p, 'Strategy & Operations   ·   FY26 planning cycle', 0.7, 4.12, 6, 0.3, {
		size: 11,
		color: c.onDarkMuted,
		spacing: 0.6,
	})
}

function northwindSummary(pptx: TsPptx, c: Palette): void {
	const p = page(pptx)
	ground(p, c.paper)
	eyebrow(p, 'Executive summary', c.muted, c.accent)
	headline(p, 'Three moves carry the plan; the rest is noise', c.ink, 27)

	const cards: readonly { value: string; unit: string; caption: string; delta: string; tint: string }[] = [
		{ value: '£412', unit: 'm', caption: 'Incremental revenue by FY29', delta: '+14% CAGR', tint: c.accent },
		{ value: '3.8', unit: 'pts', caption: 'Gross margin recovered', delta: '+180 bps in FY26', tint: c.accent2 },
		{ value: '61', unit: '%', caption: 'Of upside from two categories', delta: 'Home · Beauty', tint: c.accent3 },
	]
	cards.forEach((card, index) => {
		const x = 0.6 + index * 3.03
		plate(p, x, 1.95, 2.8, 2.25, c.mist, 0.08)
		p.slide.addShape('rect', { x, y: 1.95, w: 2.8, h: 0.075, fill: { color: card.tint }, objectName: p.id('cap') })
		p.slide.addText(
			[
				{ text: card.value, options: { fontSize: 40, bold: true, color: c.ink } },
				{ text: card.unit, options: { fontSize: 19, bold: true, color: card.tint } },
			],
			{
				x: x + 0.28,
				y: 2.24,
				w: 2.3,
				h: 0.72,
				fontFace: BODY_FACE,
				valign: 'middle',
				objectName: p.id('figure'),
			}
		)
		label(p, card.caption, x + 0.28, 3.02, 2.3, 0.62, { size: 12.5, color: c.ink })
		hairline(p, x + 0.28, 3.66, 1.0, card.tint, 1.25)
		label(p, card.delta, x + 0.28, 3.76, 2.3, 0.28, { size: 10.5, color: c.muted, bold: true, spacing: 0.8 })
	})

	label(
		p,
		'Everything below is illustrative. The point of the page is the drawing, not the numbers.',
		0.6,
		4.4,
		8.8,
		0.3,
		{ size: 10.5, color: c.muted, italic: true }
	)
	footer(p, NORTHWIND, 2, c)
}

function northwindBridge(pptx: TsPptx, c: Palette): void {
	const p = page(pptx)
	ground(p, c.paper)
	eyebrow(p, 'Revenue bridge · FY26 → FY29', c.muted, c.accent)
	headline(p, 'Category mix does the work, not price', c.ink, 27)

	// A truncated axis. Measuring from zero would draw two towers with four
	// hairlines between them — the endpoints are around £2.4bn and the movements
	// are £50–200m — so the floor is lifted and stated on the slide, which is what
	// a bridge chart does and what the note under the title says it did.
	const base = 4.4 // the chart's baseline, in slide inches
	const floor = 1900 // £m the axis starts at
	const scale = 2.35 / 1050 // inches per £m
	const bars: readonly { name: string; from: number; to: number; color: string; total?: boolean }[] = [
		{ name: 'FY26', from: floor, to: 2410, color: c.deep, total: true },
		{ name: 'Home', from: 2410, to: 2610, color: c.accent },
		{ name: 'Beauty', from: 2610, to: 2762, color: c.accent },
		{ name: 'Own brand', from: 2762, to: 2848, color: c.accent3 },
		{ name: 'Price/mix', from: 2848, to: 2796, color: 'C4405F' },
		{ name: 'FY29', from: floor, to: 2796, color: c.deep, total: true },
	]

	hairline(p, 0.62, base, 8.76, c.rule)
	bars.forEach((bar, index) => {
		const x = 0.9 + index * 1.42
		const top = base - (Math.max(bar.from, bar.to) - floor) * scale
		const height = Math.max(0.07, Math.abs(bar.to - bar.from) * scale)
		p.slide.addShape('rect', {
			x,
			y: top,
			w: 0.92,
			h: height,
			fill: { color: bar.color },
			objectName: p.id('bar'),
		})
		const delta = bar.to - bar.from
		label(
			p,
			bar.total === true ? String(bar.to) : delta >= 0 ? `+${delta}` : String(delta),
			x - 0.15,
			top - 0.3,
			1.22,
			0.26,
			{
				size: 11,
				color: bar.total === true ? c.ink : c.muted,
				bold: true,
				align: 'center',
			}
		)
		label(p, bar.name, x - 0.15, base + 0.13, 1.22, 0.26, { size: 10.5, color: c.muted, align: 'center' })
		// The connector to the next bar's floor — the line that makes a bridge read
		// as one quantity moving rather than six unrelated columns.
		if (index < bars.length - 1) {
			p.slide.addShape('line', {
				x: x + 0.92,
				y: base - (bar.to - floor) * scale,
				w: 0.5,
				h: 0,
				line: { color: c.rule, width: 1, dashType: 'dash' },
				objectName: p.id('tie'),
			})
		}
	})

	label(p, '£m, constant currency · axis truncated at 1,900', 0.62, 1.64, 5, 0.24, {
		size: 10,
		color: c.muted,
		italic: true,
	})
	footer(p, NORTHWIND, 3, c)
}

function northwindShare(pptx: TsPptx, c: Palette): void {
	const p = page(pptx)
	ground(p, c.paper)
	eyebrow(p, 'Market position', c.muted, c.accent)
	headline(p, 'We are third everywhere and first nowhere', c.ink, 27)

	const rows: readonly { name: string; parts: readonly number[] }[] = [
		{ name: 'Home', parts: [22, 19, 16, 43] },
		{ name: 'Beauty', parts: [18, 26, 14, 42] },
		{ name: 'Grocery', parts: [31, 24, 12, 33] },
		{ name: 'Apparel', parts: [14, 21, 28, 37] },
	]
	const colors = [c.accent, c.accent2, c.accent3, c.rule]
	const left = 2.0
	const width = 7.0

	rows.forEach((row, index) => {
		const y = 1.92 + index * 0.62
		label(p, row.name, 0.62, y + 0.03, 1.3, 0.3, { size: 12, color: c.ink, bold: true })
		let cursor = left
		row.parts.forEach((part, partIndex) => {
			const w = (part / 100) * width
			p.slide.addShape('rect', {
				x: cursor,
				y,
				w,
				h: 0.36,
				fill: { color: colors[partIndex] ?? c.rule },
				objectName: p.id('seg'),
			})
			if (part >= 14) {
				label(p, `${part}%`, cursor, y + 0.02, w, 0.32, {
					size: 10.5,
					color: partIndex === 3 ? c.muted : c.paper,
					bold: true,
					align: 'center',
					valign: 'middle',
				})
			}
			cursor += w
		})
	})

	const legend = ['Northwind', 'Aldergate', 'Vellum', 'All other']
	legend.forEach((name, index) => {
		const x = 2.0 + index * 1.75
		p.slide.addShape('ellipse', {
			x,
			y: 4.5,
			w: 0.16,
			h: 0.16,
			fill: { color: colors[index] ?? c.rule },
			objectName: p.id('key'),
		})
		label(p, name, x + 0.24, 4.47, 1.5, 0.26, { size: 10.5, color: c.muted })
	})
	footer(p, NORTHWIND, 4, c)
}

function northwindMatrix(pptx: TsPptx, c: Palette): void {
	const p = page(pptx)
	ground(p, c.paper)
	eyebrow(p, 'Where to play', c.muted, c.accent)
	headline(p, 'Two categories earn investment', c.ink, 27)

	// The grid is wider than it is tall, because the slide is. Bubble coordinates
	// below are fractions of the grid rather than inches, so the plot survives a
	// change to either dimension.
	const x0 = 2.35
	const y0 = 1.78
	const gw = 3.4
	const gh = 1.28
	plate(p, x0, y0, gw, gh, c.mist, 0.02)
	plate(p, x0 + gw, y0, gw, gh, 'E4E9FA', 0.02)
	plate(p, x0, y0 + gh, gw, gh, 'F6F8FD', 0.02)
	plate(p, x0 + gw, y0 + gh, gw, gh, c.mist, 0.02)
	// Only the top-right quadrant is drawn as a decision; the rest are ground.
	p.slide.addShape('rect', {
		x: x0 + gw,
		y: y0,
		w: gw,
		h: 0.055,
		fill: { color: c.accent },
		objectName: 'quadrant-cap',
	})

	hairline(p, x0, y0 + gh, gw * 2, c.rule, 1)
	p.slide.addShape('line', {
		x: x0 + gw,
		y: y0,
		w: 0,
		h: gh * 2,
		line: { color: c.rule, width: 1 },
		objectName: 'axis-v',
	})

	const bubbles: readonly { name: string; u: number; v: number; r: number; color: string }[] = [
		{ name: 'Home', u: 0.76, v: 0.3, r: 0.42, color: c.accent },
		{ name: 'Beauty', u: 0.6, v: 0.62, r: 0.34, color: c.accent },
		{ name: 'Grocery', u: 0.27, v: 0.34, r: 0.46, color: c.accent3 },
		{ name: 'Apparel', u: 0.19, v: 0.79, r: 0.28, color: c.muted },
		{ name: 'Garden', u: 0.82, v: 0.86, r: 0.22, color: c.accent2 },
	]
	for (const bubble of bubbles) {
		const cx = x0 + bubble.u * gw * 2
		const cy = y0 + bubble.v * gh * 2
		p.slide.addShape('ellipse', {
			x: cx - bubble.r,
			y: cy - bubble.r,
			w: bubble.r * 2,
			h: bubble.r * 2,
			fill: { color: bubble.color, transparency: 12 },
			objectName: p.id('bubble'),
		})
		// A label goes inside its bubble when the bubble is big enough to hold it and
		// under it when it is not. White type on a 0.44 in disc is a smear.
		const inside = bubble.r >= 0.32
		label(p, bubble.name, cx - 0.75, inside ? cy - 0.16 : cy + bubble.r + 0.03, 1.5, 0.32, {
			size: 10,
			color: inside ? c.paper : c.muted,
			bold: true,
			align: 'center',
			valign: inside ? 'middle' : 'top',
		})
	}

	label(p, 'MARKET GROWTH →', x0, y0 + gh * 2 + 0.16, gw * 2, 0.24, {
		size: 9.5,
		color: c.muted,
		align: 'center',
		spacing: 1.2,
	})
	label(p, 'RIGHT TO WIN ↑', 0.62, y0 + gh - 0.1, 1.6, 0.24, {
		size: 9.5,
		color: c.muted,
		align: 'right',
		spacing: 1.2,
	})
	label(p, 'INVEST', x0 + gw + 0.16, y0 + 0.16, 1.4, 0.24, { size: 9.5, color: c.accent, bold: true, spacing: 1.4 })
	footer(p, NORTHWIND, 5, c)
}

function northwindRoadmap(pptx: TsPptx, c: Palette): void {
	const p = page(pptx)
	ground(p, c.paper)
	eyebrow(p, 'Sequencing', c.muted, c.accent)
	headline(p, 'Four gates, eighteen months', c.ink, 27)

	const phases: readonly { name: string; when: string; note: string; color: string }[] = [
		{ name: 'Diagnose', when: 'Q1', note: 'Category P&L rebuilt bottom-up', color: c.accent3 },
		{ name: 'Commit', when: 'Q2', note: 'Two categories funded, four paused', color: c.accent2 },
		{ name: 'Build', when: 'Q3–Q4', note: 'Own-brand range live in 340 stores', color: c.accent },
		{ name: 'Scale', when: 'FY28', note: 'Operating model rewired around it', color: c.deep },
	]

	phases.forEach((phase, index) => {
		const x = 0.6 + index * 2.28
		// A chevron: five points, stated, because `chevron` is not one of the
		// presets `src/render/geometry.ts` resolves and would draw as a box.
		p.slide.addShape(ShapeType.custGeom, {
			x,
			y: 2.0,
			w: 2.26,
			h: 0.78,
			fill: { color: phase.color },
			points: [
				{ x: 0, y: 0, moveTo: true },
				{ x: 1.96, y: 0 },
				{ x: 2.26, y: 0.39 },
				{ x: 1.96, y: 0.78 },
				{ x: 0, y: 0.78 },
				{ x: 0.3, y: 0.39 },
				{ close: true },
			],
			objectName: p.id('chevron'),
		})
		label(p, phase.name, x + 0.42, 2.18, 1.6, 0.42, { size: 14, color: c.paper, bold: true, valign: 'middle' })
		label(p, phase.when.toUpperCase(), x + 0.32, 1.68, 1.8, 0.26, {
			size: 10.5,
			color: phase.color,
			bold: true,
			spacing: 1.4,
		})
		label(p, phase.note, x + 0.32, 3.0, 1.9, 0.9, { size: 11, color: c.muted })
	})

	hairline(p, 0.6, 4.16, 8.8, c.rule)
	label(p, 'Gate reviews are decision points, not status updates.', 0.6, 4.3, 6, 0.3, {
		size: 11,
		color: c.ink,
		bold: true,
	})
	footer(p, NORTHWIND, 6, c)
}

function northwindTable(pptx: TsPptx, c: Palette): void {
	const p = page(pptx)
	ground(p, c.paper)
	eyebrow(p, 'Initiative portfolio', c.muted, c.accent)
	headline(p, 'Nine initiatives, ranked by contribution', c.ink, 27)

	const head = (text: string, align: 'left' | 'right' = 'left') => ({
		text,
		options: { bold: true, color: c.paper, fill: { color: c.deep }, fontSize: 11, align, valign: 'middle' as const },
	})
	// Every cell states its fill, banded or not. A cell that states none reads back
	// as *inherited* rather than empty — `TableCell` has no `fillNoFill` accessor —
	// and the renderer paints inherited fills as a flat neutral, so an unstated
	// white row comes out grey.
	const cell =
		(band: string) =>
		(text: string, align: 'left' | 'right' = 'left', color = c.ink, bold = false) => ({
			text,
			options: { color, bold, fill: { color: band }, fontSize: 11, align, valign: 'middle' as const },
		})
	const odd = cell(c.paper)
	const even = cell(c.mist)

	p.slide.addTable(
		[
			[head('Initiative'), head('Owner'), head('Gate'), head('£m NPV', 'right'), head('Confidence', 'right')],
			[
				odd('Own-brand relaunch'),
				odd('Category'),
				odd('Q2'),
				odd('148', 'right', c.ink, true),
				odd('High', 'right', '0E9E86', true),
			],
			[
				even('Beauty concession model'),
				even('Stores'),
				even('Q2'),
				even('96', 'right', c.ink, true),
				even('High', 'right', '0E9E86', true),
			],
			[
				odd('Home range rationalisation'),
				odd('Buying'),
				odd('Q3'),
				odd('74', 'right', c.ink, true),
				odd('Medium', 'right', 'B07A12', true),
			],
			[
				even('Fulfilment network redesign'),
				even('Supply'),
				even('Q4'),
				even('61', 'right', c.ink, true),
				even('Medium', 'right', 'B07A12', true),
			],
			[
				odd('Loyalty re-platform'),
				odd('Digital'),
				odd('FY28'),
				odd('33', 'right', c.ink, true),
				odd('Low', 'right', 'C4405F', true),
			],
		],
		{
			x: 0.6,
			y: 1.9,
			w: 8.8,
			colW: [3.3, 1.5, 1.0, 1.5, 1.5],
			rowH: [0.42, 0.4, 0.4, 0.4, 0.4, 0.4],
			border: { type: 'solid', color: c.rule, width: 0.75 },
			fontFace: BODY_FACE,
			margin: 0.08,
			objectName: 'portfolio',
		}
	)
	label(p, 'Bottom four initiatives fall below the funding line and are paused, not scheduled.', 0.6, 4.42, 8.8, 0.3, {
		size: 10.5,
		color: c.muted,
		italic: true,
	})
	footer(p, NORTHWIND, 7, c)
}

function northwindClose(pptx: TsPptx, c: Palette): void {
	const p = page(pptx)
	gradientGround(p, c.deep, c.ink, 150)
	p.slide.addShape(ShapeType.custGeom, {
		x: -1.6,
		y: 2.2,
		w: 6.4,
		h: 6.4,
		fill: { color: c.accent, transparency: 70 },
		points: [...arcPoints(3.2, 3.2, 3.2, 3.2, 0, 360, true), { close: true }],
		objectName: 'orb',
	})

	label(p, 'THE ASK', 0.7, 0.9, 6, 0.28, { size: 10.5, color: c.accent3, bold: true, spacing: 2.6 })
	label(p, 'Fund two categories properly\nrather than six adequately.', 0.66, 1.4, 7.6, 1.8, {
		size: 34,
		color: c.onDark,
		bold: true,
	})
	hairline(p, 0.7, 3.5, 2.6, c.accent2, 1.25)
	const asks = [
		'£190m committed at the Q2 gate',
		'Four initiatives formally paused',
		'Category P&L owners named by March',
	]
	asks.forEach((ask, index) => {
		label(p, ask, 0.7 + index * 3.0, 3.8, 2.8, 0.7, { size: 12, color: c.onDarkMuted })
	})
}

// ---------------------------------------------------------------------------
// Deck two — Meridian Health
// ---------------------------------------------------------------------------

const MERIDIAN = 'Meridian Health · Operating model'

function meridianTitle(pptx: TsPptx, c: Palette): void {
	const p = page(pptx)
	gradientGround(p, c.ink, c.deep, 20)

	// Three concentric part-rings, off the right edge.
	const rings: readonly { outer: number; inner: number; from: number; to: number; color: string; alpha: number }[] = [
		{ outer: 3.1, inner: 2.72, from: -30, to: 190, color: c.accent, alpha: 20 },
		{ outer: 2.5, inner: 2.24, from: 40, to: 300, color: c.accent2, alpha: 34 },
		{ outer: 1.95, inner: 1.8, from: -60, to: 120, color: c.accent3, alpha: 46 },
	]
	for (const ring of rings) {
		p.slide.addShape(ShapeType.custGeom, {
			x: 6.0,
			y: -0.4,
			w: 6.4,
			h: 6.4,
			fill: { color: ring.color, transparency: ring.alpha },
			points: donutSlice(3.2, 3.2, ring.outer, ring.inner, ring.from, ring.to),
			objectName: p.id('ring'),
		})
	}

	label(p, 'BOARD PACK · WORKING DRAFT', 0.7, 0.72, 6, 0.26, {
		size: 10.5,
		color: c.onDarkMuted,
		bold: true,
		spacing: 2.6,
	})
	p.slide.addShape('rect', { x: 0.7, y: 1.5, w: 1.1, h: 0.07, fill: { color: c.accent3 }, objectName: 'kicker' })
	label(p, 'Meridian Health', 0.66, 1.78, 6, 0.95, { size: 46, color: c.onDark, bold: true })
	label(p, 'An operating model that survives the merger', 0.7, 2.78, 5.0, 0.9, { size: 17, color: c.onDarkMuted })
	hairline(p, 0.7, 3.95, 3.4, c.accent2, 1)
	label(p, 'Design phase   ·   Weeks 1–6', 0.7, 4.12, 6, 0.3, { size: 11, color: c.onDarkMuted, spacing: 0.6 })
}

function meridianPillars(pptx: TsPptx, c: Palette): void {
	const p = page(pptx)
	ground(p, c.paper)
	eyebrow(p, 'The argument', c.muted, c.accent)
	headline(p, 'Three decisions, taken in this order', c.ink, 27)

	const pillars: readonly { title: string; body: string; color: string }[] = [
		{ title: 'Decide the spans', body: 'Eleven layers between the CEO and a ward is nine too many.', color: c.accent },
		{ title: 'Move the money', body: 'Budgets follow pathways, not the buildings they inherited.', color: c.accent2 },
		{
			title: 'Fix accountability',
			body: 'One named owner per pathway, published, with the P&L attached.',
			color: c.accent3,
		},
	]

	hairline(p, 1.35, 2.62, 7.0, c.rule, 1.5)
	pillars.forEach((pillar, index) => {
		const cx = 1.35 + index * 3.5
		p.slide.addShape('ellipse', {
			x: cx - 0.46,
			y: 2.16,
			w: 0.92,
			h: 0.92,
			fill: { color: pillar.color },
			objectName: p.id('node'),
		})
		label(p, String(index + 1), cx - 0.46, 2.16, 0.92, 0.92, {
			size: 24,
			color: c.paper,
			bold: true,
			align: 'center',
			valign: 'middle',
		})
		label(p, pillar.title, cx - 1.2, 3.32, 2.4, 0.32, { size: 15, color: c.ink, bold: true, align: 'center' })
		label(p, pillar.body, cx - 1.35, 3.72, 2.7, 0.95, { size: 11.5, color: c.muted, align: 'center' })
	})
	footer(p, MERIDIAN, 2, c)
}

function meridianDonut(pptx: TsPptx, c: Palette): void {
	const p = page(pptx)
	ground(p, c.paper)
	eyebrow(p, 'Cost base · FY26', c.muted, c.accent)
	headline(p, 'Two thirds of cost sits below the line we manage', c.ink, 26)

	const slices: readonly { name: string; value: number; color: string }[] = [
		{ name: 'Clinical staff', value: 38, color: c.accent },
		{ name: 'Non-clinical staff', value: 21, color: c.accent2 },
		{ name: 'Estates', value: 17, color: c.accent3 },
		{ name: 'Supplies', value: 14, color: '2AB6A6' },
		{ name: 'Other', value: 10, color: c.rule },
	]

	let angle = 0
	for (const slice of slices) {
		const sweep = (slice.value / 100) * 360
		p.slide.addShape(ShapeType.custGeom, {
			x: 0.75,
			y: 1.62,
			w: 2.9,
			h: 2.9,
			fill: { color: slice.color },
			points: donutSlice(1.45, 1.45, 1.42, 0.86, angle + 0.9, angle + sweep - 0.9),
			objectName: p.id('slice'),
		})
		angle += sweep
	}
	label(p, '£1.9bn', 0.75, 2.72, 2.9, 0.5, { size: 24, color: c.ink, bold: true, align: 'center' })
	label(p, 'total cost base', 0.75, 3.2, 2.9, 0.26, { size: 10.5, color: c.muted, align: 'center' })

	slices.forEach((slice, index) => {
		const y = 1.78 + index * 0.52
		p.slide.addShape('rect', {
			x: 4.35,
			y: y + 0.05,
			w: 0.2,
			h: 0.2,
			fill: { color: slice.color },
			objectName: p.id('key'),
		})
		label(p, slice.name, 4.72, y, 3.0, 0.3, { size: 12, color: c.ink })
		label(p, `${slice.value}%`, 8.0, y, 1.4, 0.3, { size: 12, color: c.ink, bold: true, align: 'right' })
		hairline(p, 4.35, y + 0.4, 5.05, c.rule, 0.5)
	})
	footer(p, MERIDIAN, 3, c)
}

function meridianTrend(pptx: TsPptx, c: Palette): void {
	const p = page(pptx)
	ground(p, c.paper)
	eyebrow(p, 'Waiting list · rolling 12 months', c.muted, c.accent)
	headline(p, 'The curve turned before the merger, then stalled', c.ink, 26)

	const chart = { x: 0.7, y: 1.85, w: 8.6, h: 2.3 }
	plate(p, chart.x, chart.y, chart.w, chart.h, c.mist, 0.04)
	const series = [0.62, 0.44, 0.3, 0.52, 0.72, 1.02, 1.28, 1.2, 1.34, 1.42, 1.38, 1.46]
	const points = series.map((value, index) => ({
		x: 0.34 + (index * (chart.w - 0.68)) / (series.length - 1),
		y: chart.h - value,
	}))

	p.slide.addShape(ShapeType.custGeom, {
		x: chart.x,
		y: chart.y,
		w: chart.w,
		h: chart.h,
		fill: { color: c.accent, transparency: 82 },
		points: areaPoints(points, chart.h - 0.12),
		objectName: 'area',
	})
	p.slide.addShape(ShapeType.custGeom, {
		x: chart.x,
		y: chart.y,
		w: chart.w,
		h: chart.h,
		fill: CLEAR,
		line: { color: c.accent, width: 2.5 },
		points: points.map((point, index) => ({ x: point.x, y: point.y, moveTo: index === 0 })),
		objectName: 'trend',
	})
	for (const point of [points[5], points[11]]) {
		if (point === undefined) continue
		p.slide.addShape('ellipse', {
			x: chart.x + point.x - 0.09,
			y: chart.y + point.y - 0.09,
			w: 0.18,
			h: 0.18,
			fill: { color: c.paper },
			line: { color: c.accent, width: 2 },
			objectName: p.id('dot'),
		})
	}
	p.slide.addShape('line', {
		x: chart.x + 3.78,
		y: chart.y + 0.18,
		w: 0,
		h: chart.h - 0.36,
		line: { color: c.accent2, width: 1.5, dashType: 'dash' },
		objectName: 'merger',
	})
	label(p, 'Merger completes', chart.x + 3.9, chart.y + 0.2, 2.0, 0.28, { size: 10.5, color: c.accent2, bold: true })

	const axis = ['Apr', 'Jul', 'Oct', 'Jan', 'Apr', 'Jul']
	axis.forEach((month, index) => {
		label(p, month, chart.x + 0.1 + index * ((chart.w - 0.2) / axis.length), chart.y + chart.h + 0.1, 1.4, 0.26, {
			size: 10,
			color: c.muted,
		})
	})
	footer(p, MERIDIAN, 4, c)
}

function meridianFlow(pptx: TsPptx, c: Palette): void {
	const p = page(pptx)
	ground(p, c.paper)
	eyebrow(p, 'Referral pathway', c.muted, c.accent)
	headline(p, 'Five handoffs; three of them are queues', c.ink, 27)

	const steps: readonly { name: string; days: string; queue: boolean }[] = [
		{ name: 'Referral', days: '0d', queue: false },
		{ name: 'Triage', days: '11d', queue: true },
		{ name: 'Diagnostics', days: '24d', queue: true },
		{ name: 'Decision', days: '4d', queue: false },
		{ name: 'Treatment', days: '38d', queue: true },
	]

	steps.forEach((step, index) => {
		const x = 0.6 + index * 1.86
		plate(p, x, 2.05, 1.6, 1.05, step.queue ? c.mist : c.paper, 0.08)
		p.slide.addShape('rect', {
			x,
			y: 2.05,
			w: 1.6,
			h: 0.07,
			fill: { color: step.queue ? c.accent : '2AB6A6' },
			objectName: p.id('cap'),
		})
		if (!step.queue) {
			p.slide.addShape('roundRect', {
				x,
				y: 2.05,
				w: 1.6,
				h: 1.05,
				fill: CLEAR,
				line: { color: c.rule, width: 1 },
				rectRadius: 0.08,
				objectName: p.id('outline'),
			})
		}
		label(p, step.name, x + 0.16, 2.34, 1.3, 0.32, { size: 12.5, color: c.ink, bold: true })
		label(p, step.days, x + 0.16, 2.68, 1.3, 0.3, { size: 16, color: step.queue ? c.accent : c.muted, bold: true })
		if (index < steps.length - 1) {
			p.slide.addConnector({
				type: 'straight',
				x1: x + 1.6,
				y1: 2.58,
				x2: x + 1.86,
				y2: 2.58,
				color: c.muted,
				width: 1.25,
				endArrowType: 'triangle',
				objectName: p.id('arrow'),
			})
		}
	})

	plate(p, 0.6, 3.52, 8.8, 0.72, c.mist, 0.06)
	p.slide.addShape('rect', { x: 0.6, y: 3.52, w: 0.07, h: 0.72, fill: { color: c.accent }, objectName: 'stripe' })
	label(p, '73 of 77 days are spent waiting. The clinical work is four.', 0.85, 3.68, 8.4, 0.42, {
		size: 13.5,
		color: c.ink,
		bold: true,
		valign: 'middle',
	})
	footer(p, MERIDIAN, 5, c)
}

function meridianHeatmap(pptx: TsPptx, c: Palette): void {
	const p = page(pptx)
	ground(p, c.paper)
	eyebrow(p, 'Readiness by site', c.muted, c.accent)
	headline(p, 'Two sites are ready; four are not close', c.ink, 27)

	const heat = [c.rule, 'F7DCC9', c.accent2, c.accent] as const
	const scores: readonly (readonly [string, number, number, number, number])[] = [
		['Northgate', 3, 3, 2, 3],
		['Aldersyde', 3, 2, 3, 2],
		['Brackwell', 2, 2, 1, 1],
		['Cranholm', 1, 2, 1, 0],
		['Deepdale', 1, 1, 0, 1],
		['Elverton', 0, 1, 0, 0],
	]
	const words = ['Not started', 'Early', 'Progressing', 'Ready']
	// Stated fills for the same reason as the portfolio table: an unstated cell is
	// an inherited one, and inherited paints grey.
	const head = (text: string, align: 'left' | 'center' = 'center') => ({
		text,
		options: {
			bold: true,
			color: c.muted,
			fill: { color: c.paper },
			fontSize: 10.5,
			align,
			valign: 'middle' as const,
		},
	})

	p.slide.addTable(
		[
			[head('SITE', 'left'), head('GOVERNANCE'), head('DATA'), head('WORKFORCE'), head('FINANCE')],
			...scores.map(([site, ...values]) => [
				{
					text: site,
					options: { bold: true, color: c.ink, fill: { color: c.mist }, fontSize: 11.5, valign: 'middle' as const },
				},
				...values.map((value) => ({
					text: words[value] ?? '',
					options: {
						fill: { color: heat[value] ?? c.rule },
						color: value >= 2 ? c.paper : c.muted,
						bold: value >= 2,
						fontSize: 10.5,
						align: 'center' as const,
						valign: 'middle' as const,
					},
				})),
			]),
		],
		{
			x: 0.6,
			y: 1.86,
			w: 8.8,
			colW: [2.0, 1.7, 1.7, 1.7, 1.7],
			rowH: 0.38,
			border: { type: 'solid', color: c.paper, width: 2 },
			fontFace: BODY_FACE,
			margin: 0.06,
			objectName: 'readiness',
		}
	)
	footer(p, MERIDIAN, 6, c)
}

function meridianStatement(pptx: TsPptx, c: Palette): void {
	const p = page(pptx)
	ground(p, c.mist)
	p.slide.addShape(ShapeType.custGeom, {
		x: -2.2,
		y: -2.2,
		w: 5.6,
		h: 5.6,
		fill: { color: c.accent3, transparency: 62 },
		points: [...arcPoints(2.8, 2.8, 2.8, 2.8, 0, 360, true), { close: true }],
		objectName: 'orb',
	})
	p.slide.addShape('rect', { x: 0, y: H - 0.16, w: W, h: 0.16, fill: { color: c.accent }, objectName: 'foot' })

	label(p, '“', 0.62, 0.9, 1.6, 1.4, { size: 96, color: c.accent, bold: true, face: 'Georgia' })
	label(p, 'The structure is not the problem.\nNobody can say who decides.', 1.5, 1.55, 7.6, 2.0, {
		size: 30,
		color: c.ink,
		bold: true,
		face: 'Georgia',
	})
	hairline(p, 1.55, 3.72, 1.6, c.accent, 1.5)
	label(p, 'Divisional director, week 2 interviews', 1.55, 3.9, 6, 0.3, { size: 11.5, color: c.muted, italic: true })
}

function meridianNext(pptx: TsPptx, c: Palette): void {
	const p = page(pptx)
	gradientGround(p, c.deep, c.ink, 160)

	label(p, 'NEXT SIX WEEKS', 0.7, 0.78, 6, 0.28, { size: 10.5, color: c.accent3, bold: true, spacing: 2.6 })
	label(p, 'What we need from the board', 0.66, 1.18, 8, 0.7, { size: 32, color: c.onDark, bold: true })
	hairline(p, 0.7, 2.05, 2.6, c.accent2, 1.25)

	const steps: readonly { n: string; text: string }[] = [
		{ n: '01', text: 'Sign off the three design principles as written' },
		{ n: '02', text: 'Name the pathway owners — six people, by 14 March' },
		{ n: '03', text: 'Release the transition budget against the Q2 gate' },
	]
	steps.forEach((step, index) => {
		const y = 2.35 + index * 0.78
		label(p, step.n, 0.7, y, 0.8, 0.4, { size: 20, color: c.accent2, bold: true })
		label(p, step.text, 1.6, y + 0.04, 7.4, 0.4, { size: 15, color: c.onDark })
		hairline(p, 0.7, y + 0.56, 8.6, '3A2246', 0.75)
	})
	label(p, 'Meridian Health · Operating model · Draft for discussion', 0.7, H - 0.62, 8.6, 0.3, {
		size: 10,
		color: c.onDarkMuted,
		spacing: 0.8,
	})
}

// ---------------------------------------------------------------------------
// The decks
// ---------------------------------------------------------------------------

export interface ShowcaseDeck {
	/** Stable id, used as the marquee row's key. */
	name: string
	/** What the row is called on the page. */
	title: string
	/** One line under the row. */
	blurb: string
	/** What the row's download is called on disk. Stated, not derived from `title`. */
	file: string
	build: () => Promise<Uint8Array>
}

/** A writer with fixed properties: no clock, no host, nothing that drifts per visit. */
function newDeck(title: string): TsPptx {
	const pptx = new TsPptx()
	pptx.layout = 'LAYOUT_16x9'
	pptx.author = 'pptx-html'
	pptx.company = 'pptx-html showcase'
	pptx.subject = 'Illustrative deck, written by @shbernal/ts-pptx'
	pptx.title = title
	return pptx
}

async function bytesOf(pptx: TsPptx): Promise<Uint8Array> {
	return (await pptx.write({ outputType: 'uint8array' })) as Uint8Array
}

export const SHOWCASE: readonly ShowcaseDeck[] = [
	{
		name: 'northwind',
		title: 'Northwind Retail — Growth review',
		blurb: 'Eight slides: waterfall, stacked share, a 2×2, a chevron roadmap and a ranked portfolio table.',
		file: 'northwind-retail-growth-review.pptx',
		build: async () => {
			const pptx = newDeck('Northwind Retail — Growth review')
			northwindTitle(pptx, INDIGO)
			northwindSummary(pptx, INDIGO)
			northwindBridge(pptx, INDIGO)
			northwindShare(pptx, INDIGO)
			northwindMatrix(pptx, INDIGO)
			northwindRoadmap(pptx, INDIGO)
			northwindTable(pptx, INDIGO)
			northwindClose(pptx, INDIGO)
			return bytesOf(pptx)
		},
	},
	{
		name: 'meridian',
		title: 'Meridian Health — Operating model',
		blurb: 'Eight slides: concentric rings, a donut, an area trend, a pathway flow and a per-cell heatmap.',
		file: 'meridian-health-operating-model.pptx',
		build: async () => {
			const pptx = newDeck('Meridian Health — Operating model')
			meridianTitle(pptx, PLUM)
			meridianPillars(pptx, PLUM)
			meridianDonut(pptx, PLUM)
			meridianTrend(pptx, PLUM)
			meridianFlow(pptx, PLUM)
			meridianHeatmap(pptx, PLUM)
			meridianStatement(pptx, PLUM)
			meridianNext(pptx, PLUM)
			return bytesOf(pptx)
		},
	},
]
