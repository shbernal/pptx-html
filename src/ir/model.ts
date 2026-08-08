/**
 * Intermediate representation (IR) — the contract between the browser-only
 * `extract/` layer (DOM → IR) and the pure `emit/` layer (IR → ts-pptx).
 *
 * Positions are in **inches** (slide-space). This file is the single source of
 * truth for the IR shape; both layers depend on it and nothing else.
 *
 * The text / list / table item shapes are deliberately ts-pptx-flavoured: runs
 * and table cells are `{ text, options }` objects passed straight to ts-pptx's
 * `addText` / `addTable`, and text styling lives in a `style` object the emit
 * layer spreads into text options. This mirrors what `extract/extractor.ts`
 * actually produces, which is what lets `emit/slide.ts` be type-checked against
 * the IR rather than suppressed with `@ts-nocheck`.
 */

/** A rectangle in slide-space, inches. */
export interface Rect {
	x: number
	y: number
	w: number
	h: number
}

/**
 * A run of text within a text/list item, in ts-pptx's `addText` shape. `options`
 * is passed through opaquely (bold/italic/color/bullet/hyperlink/breakLine/…), so
 * it is intentionally loose.
 */
export interface PptxRun {
	text: string
	options?: Record<string, unknown>
}

/** Inline text styling, spread into text-box options by the emit layer. */
export interface TextStyle {
	fontFace?: string
	fontSize?: number
	color?: string
	bold?: boolean
	italic?: boolean
	underline?: boolean
	strike?: boolean
	align?: 'left' | 'center' | 'right' | 'justify'
	valign?: 'top' | 'middle' | 'bottom'
	/** Text-box inset, points (or `[top, right, bottom, left]`). */
	margin?: number | number[]
	lineSpacing?: number
	/** 0–100, where 100 is fully transparent (ts-pptx convention). */
	transparency?: number
	/** Size the box to the layout rect instead of the text-fit heuristics. */
	constrainTextBox?: boolean
}

/** A stroke (border / line / table border). */
export interface Line {
	color?: string
	/** Stroke width in points. */
	width?: number
	dash?: 'solid' | 'dash' | 'dot' | 'dashDot' | 'lgDash' | 'sysDash' | 'sysDot'
}

/**
 * A table / cell border. Passed through to ts-pptx opaquely; the extractor may
 * emit a single border or an array of per-side borders, so the shape is loose.
 */
export type Border = Record<string, unknown>

/** A single table cell, in ts-pptx's `addTable` shape (`{ text, options }`). */
export interface TableCell {
	text: string
	options?: Record<string, unknown>
}

/** A point on a freeform (custGeom) path. Mirrors ts-pptx's freeform points. */
export type FreeformPoint =
	| { x: number; y: number; moveTo?: boolean }
	| { x: number; y: number; curve: { type: 'cubic'; x1: number; y1: number; x2: number; y2: number } }
	| { x: number; y: number; curve: { type: 'quadratic'; x1: number; y1: number } }
	| { x: number; y: number; curve: { type: 'arc'; hR: number; wR: number; stAng: number; swAng: number } }
	| { close: true }

interface BaseItem {
	position: Rect
	/** z-order; higher draws on top. */
	z: number
	/** Human-readable label for warnings/debugging (e.g. "table cell overlay 1.2"). */
	label?: string
	/** Marks purely decorative imagery/shapes (backgrounds, ornaments). */
	kind?: 'decor'
}

export interface TextItem extends BaseItem {
	type: 'text' | 'list'
	/** Plain string, or ts-pptx runs for rich / bulleted text. */
	text: string | PptxRun[]
	/** Originating tag (e.g. `h1`, `p`, `span`); drives heading text-fit heuristics. */
	tag?: string
	/** Disables wrapping; widens the text-fit box. */
	noWrap?: boolean
	style?: TextStyle
}

export interface ShapeItem extends BaseItem {
	type: 'shape'
	/** `null` when the shape has no fill (extractor emits `null` for unfilled). */
	fill?: string | null
	line?: Line
	/** 0–100, where 100 is fully transparent (ts-pptx convention). */
	transparency?: number
	/** Corner radius in inches, for rounded rectangles. */
	radius?: number
}

export interface ImageItem extends BaseItem {
	type: 'image'
	src: string
	objectFit?: 'cover' | 'contain' | 'fill'
	/** 0–100, where 100 is fully transparent (ts-pptx convention). */
	transparency?: number
}

export interface LineItem extends BaseItem {
	type: 'line'
	color?: string
	/** Stroke width in points. */
	width?: number
	dash?: Line['dash']
}

export interface TableItem extends BaseItem {
	type: 'table'
	rows: TableCell[][]
	/** Column widths in inches. */
	colW?: number[]
	/** Row heights in inches. */
	rowH?: number[]
	border?: Border
}

/** Freeform / custGeom path. Emitted by `emit/custgeom.ts`. */
export interface PathItem extends BaseItem {
	type: 'path'
	points: FreeformPoint[]
	fill?: string
	line?: Line
}

export type Item = TextItem | ShapeItem | ImageItem | LineItem | TableItem | PathItem

export type Background =
	| { type: 'color'; value: string }
	| { type: 'image'; src: string; fallback?: string }

/** Slide dimensions in inches. */
export interface SlideSize {
	width: number
	height: number
}

export interface SlideModel {
	background: Background
	/** z-ordered items on the slide. */
	items: Item[]
	notes?: string
}
