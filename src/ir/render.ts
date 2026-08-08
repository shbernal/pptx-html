/**
 * `RenderIr` — the **paint** model: a typed, drawable representation of a deck
 * that an HTML/SVG document is rendered *from*.
 *
 * ## What this is not
 *
 * It is **not** the round-trip contract and **not** the emit spec. `DeckIr` from
 * `@shbernal/ts-pptx/script` is both of those, and `diffDeckIr` judges it. The
 * two exist side by side for one reason: `DeckIr.slides[].calls[].args` is
 * `IrValue` — untyped write-API option bags — and you cannot draw an SVG from a
 * bag. So:
 *
 * - **emit + verify** speak `DeckIr`. One round-trip contract, upstream-owned.
 * - **render** speaks `RenderIr`. Typed for geometry and paint, mapped from the
 *   read model directly.
 *
 * Both are built from the same read-model traversal, so they cannot disagree
 * about what the source deck contained. Merging them would either drag geometry
 * into upstream's write-call vocabulary or drag write-API arguments into the
 * renderer — and the second is how a renderer starts silently deciding what gets
 * emitted. **Nothing in `src/emit/` may import this file.**
 *
 * ## Units
 *
 * **EMU everywhere.** Every length that describes a position or a size is an
 * integer count of English Metric Units (914 400 per inch); point-valued fields
 * — stroke widths, font sizes, text insets — say `Pt` in their name because
 * OOXML states them that way and rounding them to EMU would be lossy in the
 * other direction. Inches appear only at the edges, via {@link inchesOf} /
 * {@link emuOf}.
 *
 * A deck read from a package has exact integer EMU on every frame. Storing
 * inches instead — or, worse, storing both — reintroduces float noise on every
 * element, which then shows up as a spurious difference on every diff.
 *
 * ## JSON is the wire format
 *
 * Part 05 embeds this model in the rendered document as a JSON island and part
 * 06 parses it back, so every type here must survive `JSON.parse(JSON.stringify(…))`
 * unchanged. Concretely: no `undefined` values (an absent field is a missing
 * key, which is the single spelling of "absent"), no `Date`, no `Map`/`Set`, and
 * **no `Uint8Array`** — bytes live behind an {@link AssetRef}, never inline.
 */

import type { AssetRef, FidelityNote, SlideLayoutIr, SlideSource } from '@shbernal/ts-pptx/script'
// `SlideSource` is two unrelated types upstream: `'authored' | 'carried'` here in
// `/script`, and the `{ extractSlides() }` interface in `/read`. Importing both
// into one file silently shadows one of them.
import type { ColorTransform, GeometryCommand, LineSpacing, ThemeColorSlot } from '@shbernal/ts-pptx/read'

export type { AssetRef, FidelityNote, GeometryCommand, SlideSource }

/**
 * The shape version of this IR. Part 05's HTML island carries it and part 06's
 * parser refuses a mismatch loudly rather than mis-reading a document written by
 * an older renderer. Bump on any change that is not purely additive-optional.
 *
 * `1` was the legacy DOM-shaped IR in `./model`.
 */
export const IR_VERSION = 2

/** English Metric Units per inch. The one conversion constant in the model. */
export const EMU_PER_INCH = 914_400

/** EMU → inches. For the legacy boundary and for human-readable output only. */
export function inchesOf(emu: number): number {
	return emu / EMU_PER_INCH
}

/** Inches → EMU, rounded to the integer the format actually stores. */
export function emuOf(inches: number): number {
	return Math.round(inches * EMU_PER_INCH)
}

// ---------------------------------------------------------------------------
// Identity
// ---------------------------------------------------------------------------

/**
 * A node's identity, stable across the whole loop: import assigns it, the
 * renderer writes it into the document, the return path reads it back, and the
 * differ aligns the two sides by it instead of by position.
 *
 * Ids are **derived from the source, never generated**, so re-importing the same
 * deck twice yields the same ids — which is what lets a second import be
 * compared against the first at all. See {@link importedNodeId} /
 * {@link domNodeId} for the two derivations.
 */
export type NodeId = string

/**
 * Identity for a node imported from a package: source slide number (1-based) and
 * the shape's `p:cNvPr/@id`, which is unique within a slide's shape tree.
 *
 * Structural rather than hashed. A hash would buy opacity and cost collisions;
 * the pair is already unique, already stable, and reads back in a stack trace.
 */
export function importedNodeId(slideNumber: number, shapeId: number): NodeId {
	return `s${slideNumber}.sp${shapeId}`
}

/** Identity for a table cell, which has no `cNvPr` of its own. */
export function cellNodeId(table: NodeId, row: number, column: number): NodeId {
	return `${table}.r${row}c${column}`
}

/**
 * Identity for a node extracted from the DOM, where there is no `cNvPr/@id`:
 * the child-index path from the slide root. Deterministic for a given document,
 * which is all the inference lane can offer — a DOM edit that inserts an element
 * renumbers its siblings, and that is correctly read as drift.
 */
export function domNodeId(slideNumber: number, childPath: readonly number[]): NodeId {
	return `s${slideNumber}.dom${childPath.length > 0 ? `-${childPath.join('-')}` : ''}`
}

// ---------------------------------------------------------------------------
// Geometry and placement
// ---------------------------------------------------------------------------

/** A rectangle in slide space, EMU. */
export interface Box {
	x: number
	y: number
	w: number
	h: number
}

/** Per-edge fractions (`a:srcRect` / `a:fillRect`); `0.1` is 10 %, negatives bleed. */
export interface EdgeRect {
	left: number
	top: number
	right: number
	bottom: number
}

/**
 * Where and how a node sits: its box plus the `a:xfrm` attributes the legacy IR
 * had no representation for at all.
 */
export interface Placement {
	box: Box
	/** `a:xfrm/@rot`, clockwise degrees. */
	rotation: number
	flipH: boolean
	flipV: boolean
}

/**
 * A shape's outline, either a named preset or a freeform path list.
 *
 * `custom` reuses the read model's own {@link GeometryCommand} vocabulary rather
 * than a private one, so an imported freeform maps across one-to-one and
 * `src/emit/custgeom.ts` keeps passing the same verbs to the writer's freeform
 * DSL. Coordinates are raw path units in each path's own `0..w`/`0..h` space —
 * not EMU — exactly as the read model reports them.
 */
export type Geometry =
	| {
			kind: 'preset'
			/** `a:prstGeom/@prst`, e.g. `roundRect`. */
			preset: string
			/** `a:avLst` guide values, keyed by name (`adj`, `adj1`, …). Open-ended by the schema. */
			adjustValues: Record<string, number>
	  }
	| { kind: 'custom'; paths: GeometryPath[] }

/** One `<a:path>` of a freeform geometry. Order of `commands` *is* the geometry. */
export interface GeometryPath {
	/** Path-unit width (`@w`); the denominator for this path's `x` coordinates. */
	w: number
	/** Path-unit height (`@h`); the denominator for this path's `y` coordinates. */
	h: number
	/** `@fill` (`norm`/`none`/`lighten`/…). */
	fill: string
	/** `@stroke`. */
	stroke: boolean
	commands: GeometryCommand[]
}

// ---------------------------------------------------------------------------
// Paint
// ---------------------------------------------------------------------------

/**
 * A colour, keeping the theme reference intact.
 *
 * The legacy IR stored a hex string, which throws away both the `schemeClr`
 * token and its transform list — precisely the information needed to show that a
 * theme reference survived the loop rather than being flattened to the colour it
 * happened to resolve to on import. `effectiveHex` is what a renderer paints;
 * `slot` + `transforms` is what the deck actually said.
 */
export type Color =
	| { kind: 'srgb'; hex: string; alpha?: number }
	| {
			kind: 'scheme'
			slot: ThemeColorSlot
			/** `lumMod`/`lumOff`/`shade`/`tint`/… in document order, as read. */
			transforms: ColorTransform[]
			/** `slot` resolved through the theme with `transforms` applied. */
			effectiveHex: string
			alpha?: number
	  }

/** One stop of a gradient, position 0–1. */
export interface GradientStop {
	position: number
	color: Color
}

export type Gradient =
	| { kind: 'linear'; angleDeg: number; stops: GradientStop[] }
	| { kind: 'path'; shape: 'circle' | 'rect' | 'shape'; stops: GradientStop[] }

/**
 * A fill. `none` is a state of its own rather than an absent value — the legacy
 * IR overloaded `null` to mean both "no fill" and "not stated", which are
 * different decks.
 */
export type Fill =
	| { kind: 'none' }
	| { kind: 'solid'; color: Color }
	| { kind: 'gradient'; gradient: Gradient }
	| { kind: 'pattern'; preset: string | null; foreground: Color | null; background: Color | null }
	| {
			kind: 'picture'
			asset: AssetRef
			mode: 'stretch' | 'tile'
			/** Source crop (`a:srcRect`). */
			srcRect?: EdgeRect
			alpha?: number
	  }

export type DashStyle =
	| 'solid'
	| 'dot'
	| 'dash'
	| 'lgDash'
	| 'dashDot'
	| 'lgDashDot'
	| 'lgDashDotDot'
	| 'sysDash'
	| 'sysDot'
	| 'sysDashDot'
	| 'sysDashDotDot'

/** An arrowhead (`a:headEnd` / `a:tailEnd`). */
export interface LineEnd {
	type: 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow'
	width: 'sm' | 'med' | 'lg'
	length: 'sm' | 'med' | 'lg'
}

/** An outline: shape border, connector line, or table cell edge. */
export interface Stroke {
	widthPt: number
	color: Color
	dash: DashStyle
	cap: 'flat' | 'round' | 'square'
	join: 'round' | 'bevel' | 'miter'
	head?: LineEnd
	tail?: LineEnd
}

// ---------------------------------------------------------------------------
// Text
// ---------------------------------------------------------------------------

/**
 * Character formatting on a run.
 *
 * Every field is optional and an absent one means *inherited*, not *default* —
 * the distinction the legacy `TextStyle` could not make. Import resolves the
 * inheritance chain for rendering but records what the run itself stated, so
 * emit does not bake a placeholder's inherited size into the slide.
 */
export interface RunProperties {
	fontFace?: string
	sizePt?: number
	bold?: boolean
	italic?: boolean
	/**
	 * `a:rPr/@u`. `ST_TextUnderlineType` has seventeen members; the three modeled
	 * here are the ones the write API can express. Anything else must be imported
	 * as a {@link FidelityNote}, not rounded to `single`.
	 */
	underline?: 'none' | 'single' | 'double'
	strike?: 'none' | 'single' | 'double'
	color?: Color
	/** `a:rPr/@spc`, character spacing. */
	spacingPt?: number
	/** Superscript/subscript (`@baseline`), as a percentage of font size. */
	baselinePct?: number
	hyperlink?: Hyperlink
}

/** A run's click hyperlink (`a:hlinkClick`). */
export interface Hyperlink {
	/** External URL, or `null` for an internal slide jump. */
	url: string | null
	/** 1-based target slide, for an internal jump. */
	targetSlide?: number
	tooltip?: string
}

/** A paragraph's bullet definition (`a:buNone` / `a:buChar` / `a:buAutoNum`). */
export type Bullet =
	| { kind: 'none' }
	| { kind: 'character'; char: string; font?: string; color?: Color; sizePct?: number }
	| {
			kind: 'number'
			/** `a:buAutoNum/@type`, e.g. `arabicPeriod`. */
			scheme: string
			startAt: number
			font?: string
			color?: Color
			sizePct?: number
	  }

export interface ParagraphProperties {
	align: 'left' | 'center' | 'right' | 'justify'
	/** `a:pPr/@lvl`, 0-based outline depth. */
	level: number
	bullet: Bullet
	/** `@marL` / `@indent`, points. */
	marginLeftPt?: number
	indentPt?: number
	/** `a:lnSpc`, in whichever of the two OOXML forms the source used. */
	lineSpacing?: LineSpacing
	spaceBeforePt?: number
	spaceAfterPt?: number
}

export interface TextRun {
	text: string
	props: RunProperties
}

export interface Paragraph {
	props: ParagraphProperties
	runs: TextRun[]
}

/**
 * A text frame: the paragraphs plus the `a:bodyPr` state that decides how they
 * are laid out inside the box. Autofit in particular has to be modeled rather
 * than recomputed — a renderer that re-measures text would make emit
 * machine-dependent, which is the one thing the charter rules out.
 */
export interface TextBody {
	paragraphs: Paragraph[]
	/** Write-side spelling of `AutofitMode`: `none` / `normAutofit` / `spAutoFit`. */
	autofit: 'none' | 'shrink' | 'resize'
	anchor: 'top' | 'middle' | 'bottom'
	wrap: boolean
	/** `@lIns`/`@rIns`/`@tIns`/`@bIns`, points, defaults already resolved. */
	insetsPt: { left: number; right: number; top: number; bottom: number }
	/** `@vert`, when the frame is not horizontal. */
	vertical?: 'vert' | 'vert270' | 'eaVert' | 'wordArtVert'
}

// ---------------------------------------------------------------------------
// Nodes
// ---------------------------------------------------------------------------

/** Which placeholder on the layout a node inherits from. */
export interface PlaceholderLink {
	/** `p:ph/@type`, e.g. `title`, `body`; `null` for the unnamed default. */
	type: string | null
	/** `p:ph/@idx`. */
	idx: string
}

interface NodeBase {
	id: NodeId
	/** `p:cNvPr/@name`. Carried because {@link FidelityNote.shapeName} points at it. */
	name: string
	placement: Placement
	/**
	 * How this node is drawn — a **rendering** fact, not a fidelity one. A
	 * `placeholder` node renders as a labelled inert box naming
	 * {@link NodeBase.standsFor} and nothing more; there is no preview, because
	 * `.pptx` → PNG cannot run in a browser and a node with no model is exactly
	 * the case with nothing to draw from.
	 *
	 * Fidelity is recorded separately, as {@link RenderSlide.fidelity}.
	 */
	render: 'drawn' | 'placeholder'
	/** The construct a `placeholder` node stands in for, e.g. `chart`, `smartArt`. */
	standsFor?: string
	/** `p:cNvPr/@hidden`. */
	hidden?: boolean
	/** `p:cNvPr/@descr`, alt text. */
	alt?: string
	/** Set when the node is a layout placeholder rather than a free shape. */
	placeholder?: PlaceholderLink
}

export interface ShapeNode extends NodeBase {
	kind: 'shape'
	geometry: Geometry
	fill: Fill
	stroke: Stroke | null
	text: TextBody | null
}

export interface PictureNode extends NodeBase {
	kind: 'picture'
	asset: AssetRef
	/** `a:srcRect` crop. */
	crop?: EdgeRect
	geometry: Geometry
	stroke: Stroke | null
}

/** One bound end of a connector. */
export interface Connection {
	/** The bound node, by {@link NodeId}. `null` for a dangling binding. */
	node: NodeId | null
	/** Connection-site index on the bound shape. */
	site: number
}

export interface ConnectorNode extends NodeBase {
	kind: 'connector'
	geometry: Geometry
	stroke: Stroke | null
	start?: Connection
	end?: Connection
}

export interface TableColumn {
	widthEmu: number
}

export interface TableCell {
	id: NodeId
	text: TextBody | null
	fill: Fill
	borders: {
		left: Stroke | null
		right: Stroke | null
		top: Stroke | null
		bottom: Stroke | null
	}
	/** `>1` when this cell spans; `null` for a plain 1×1 cell. */
	span: { columns: number; rows: number } | null
	/** `true` for a cell covered by a neighbour's span (`hMerge`/`vMerge`). */
	covered: boolean
	marginsPt: { left: number; right: number; top: number; bottom: number }
	anchor: 'top' | 'middle' | 'bottom'
}

export interface TableRow {
	heightEmu: number
	cells: TableCell[]
}

export interface TableNode extends NodeBase {
	kind: 'table'
	columns: TableColumn[]
	rows: TableRow[]
}

/**
 * A group. Z-order is tree order — the flat `z` number the legacy IR carried
 * cannot express "inside this group, above that sibling".
 */
export interface GroupNode extends NodeBase {
	kind: 'group'
	/** `a:grpSpPr/a:xfrm/a:chOff`: the origin children's coordinates are relative to. */
	childOffset: { x: number; y: number }
	/** `a:chExt`: the extent that child space is scaled from, onto {@link Placement.box}. */
	childExtent: { w: number; h: number }
	children: RenderNode[]
}

/**
 * A node whose construct this IR does not model — a chart, a SmartArt diagram,
 * an embedded object. It is *not* a loss: the slide carrying it is `carried`, so
 * the original XML crosses intact. This node exists only so the renderer has
 * something to put on the page, and it always renders as a placeholder.
 */
export interface OpaqueNode extends NodeBase {
	kind: 'opaque'
	render: 'placeholder'
	standsFor: string
}

export type RenderNode = ShapeNode | PictureNode | ConnectorNode | TableNode | GroupNode | OpaqueNode

/** Every discriminator of {@link RenderNode}. Enumerated so tests can assert coverage. */
export const NODE_KINDS = ['shape', 'picture', 'connector', 'table', 'group', 'opaque'] as const

export type NodeKind = (typeof NODE_KINDS)[number]

// ---------------------------------------------------------------------------
// Slides and deck
// ---------------------------------------------------------------------------

export type SlideBackground =
	| { kind: 'none' }
	| { kind: 'solid'; color: Color }
	| { kind: 'gradient'; gradient: Gradient }
	| { kind: 'picture'; asset: AssetRef }

export interface Background {
	/** Which tier of slide → layout → master the background resolved from. */
	source: 'slide' | 'layout' | 'master'
	fill: SlideBackground
}

/**
 * The residual channel: source XML kept verbatim for anything not modeled.
 *
 * Attached at **slide granularity** in v1 — the simplest thing that is provably
 * correct, and the one shipped upstream API supports it end to end
 * (`ExtractedSlide` carries a standalone `<p:sld>` plus its media, and
 * `importSlide` / `appendSlides` reproduce the rel graph). Shape-granularity
 * residual, mixing modeled and carried elements on one slide, depends on
 * `importShape` behaving well enough and is deliberately not depended on here.
 */
export interface Residual {
	kind: 'slide' | 'shape'
	/** The standalone part body, verbatim. */
	xml: string
	/** Media the XML references, by {@link AssetRef} — never inline bytes. */
	assets: AssetRef[]
}

export interface RenderSlide {
	/** 1-based index in the source deck — the identity {@link FidelityNote.slideNumber} points at. */
	number: number
	/**
	 * `authored` (transcribed into write-API calls) or `carried` (holds a
	 * construct the write API cannot express, so the slide is copied from the
	 * source package instead). Upstream's framing: `carried` is a
	 * *recommendation*, not an erasure — the nodes are populated either way, so
	 * the renderer always has something to draw.
	 */
	source: SlideSource
	/** The layout to bind to, or `null` when the source slide resolves none. */
	layout: SlideLayoutIr | null
	hidden: boolean
	background: Background
	/** Paint order is array order, front-to-back, matching OOXML document order. */
	nodes: RenderNode[]
	/** Speaker notes as plain text; notes-slide geometry does not survive. */
	notes: string | null
	residual: Residual | null
	/**
	 * Declared losses for this slide, in upstream's vocabulary
	 * (`Disposition` × `Cause`). **Do not coin a local
	 * modeled/carried/unsupported enum alongside it**: two vocabularies for one
	 * concept is how the differ and the renderer drift apart.
	 */
	fidelity: FidelityNote[]
}

/** An asset's identity in the manifest. Bytes are *not* here — see the file header. */
export interface AssetManifestEntry {
	/** The key an {@link AssetRef.$asset} resolves against. */
	name: string
	contentType: string
	byteLength: number
	/** Lowercase hex SHA-256 of the bytes. Media is addressed by content, never re-embedded. */
	sha256: string
}

export interface RenderIr {
	irVersion: number
	/** Slide dimensions, EMU. */
	size: { w: number; h: number }
	slides: RenderSlide[]
	/**
	 * Every asset any slide references, by name. The island carries this manifest
	 * and never the bytes; whether the bytes ride along in the document is a
	 * *render* option, not an IR one.
	 */
	assets: AssetManifestEntry[]
}
