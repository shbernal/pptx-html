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
 * **EMU everywhere a position or a size is stated** — an integer count of
 * English Metric Units, 914 400 per inch. Where a value arrives in points
 * instead (stroke widths, font sizes, the text insets the read model already
 * converts), the field name says so with a `Pt` suffix, and where it arrives in
 * EMU the name says `Emu`. **Nothing here converts a unit**: a field carries the
 * unit its source reported, so no rounding decision is taken at this layer.
 * Inches appear only at the edges, via {@link inchesOf} / {@link emuOf}.
 *
 * A deck read from a package has exact integer EMU on every frame. Storing
 * inches instead — or, worse, storing both — reintroduces float noise on every
 * element, which then shows up as a spurious difference on every diff.
 *
 * ## JSON is the wire format
 *
 * `src/render/` embeds this model in the rendered document as a JSON island and
 * `src/parse/` reads it back, so every type here must survive
 * `JSON.parse(JSON.stringify(…))` unchanged. Concretely: no `undefined` values (an absent field is a missing
 * key, which is the single spelling of "absent"), no `Date`, no `Map`/`Set`, and
 * **no `Uint8Array`** — bytes live behind an {@link AssetRef}, never inline.
 */

import type { AssetRef, FidelityNote, SlideLayoutIr, SlideSource } from '@shbernal/ts-pptx/script'
// `SlideSource` is two unrelated types upstream: `'authored' | 'carried'` here in
// `/script`, and the `{ extractSlides() }` interface in `/read`. Importing both
// into one file silently shadows one of them.
import type { ColorMapToken, ColorTransform, GeometryCommand, LineSpacing, ThemeColorSlot } from '@shbernal/ts-pptx/read'

export type { AssetRef, FidelityNote, GeometryCommand, SlideSource }

/**
 * The shape version of this IR. The HTML island carries it and the parser
 * refuses a mismatch loudly rather than mis-reading a document written by an
 * older renderer. Bump on any change that is not purely additive-optional.
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
 *
 * The box is **slide-absolute**, with every enclosing group transform already
 * composed in (the read model's `absoluteFrame` does this). A group child's own
 * `a:xfrm` is stated in its group's child space (`a:chOff`/`a:chExt`) and is not
 * directly placeable, so composing it here rather than at paint time is what
 * keeps a nested group from being subtly displaced — and it is why
 * {@link GroupNode} carries no child-space transform of its own.
 */
export interface Placement {
	box: Box
	/** `a:xfrm/@rot`, clockwise degrees, after composing enclosing group rotations. */
	rotation: number
	flipH: boolean
	flipV: boolean
	/**
	 * Which tier of slide → layout → master the box resolved from.
	 *
	 * This is the "model the link, do not flatten it" half of placeholder
	 * inheritance: a renderer needs a concrete box, but a box silently promoted
	 * from `layout` to `own` is a slide that has stopped tracking its layout, and
	 * the differ would see it. Recording where it came from keeps both true.
	 */
	geometrySource: 'own' | 'layout' | 'master'
}

/**
 * A shape's outline, either a named preset or a freeform path list.
 *
 * `custom` reuses the read model's own {@link GeometryCommand} vocabulary rather
 * than a private one, so an imported freeform maps across one-to-one and
 * `src/heuristic/custgeom.ts` keeps passing the same verbs to the writer's freeform
 * DSL. Coordinates are raw path units in each path's own `0..w`/`0..h` space —
 * not EMU — exactly as the read model reports them.
 */
export type Geometry =
	| {
			kind: 'preset'
			/** `a:prstGeom/@prst`, e.g. `roundRect`. */
			preset: string
			/**
			 * `a:avLst` guide values, keyed by name (`adj`, `adj1`, …), as the raw
			 * `a:gd/@fmla` the source states — `'val 13333'`, or a computed form whose
			 * verb is an operator name. Strings rather than numbers because the
			 * schema's guide grammar is a formula language, and parsing `val N` while
			 * silently dropping everything else would turn an unmodeled adjust handle
			 * into an unnoticed shape change.
			 */
			adjustValues: Record<string, string>
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
/**
 * What an `a:schemeClr/@val` may name.
 *
 * Not {@link ThemeColorSlot} alone: inside a slide the reference usually goes
 * through the colour *map* (`bg1`/`tx1`/`bg2`/`tx2`), which points at a slot
 * rather than being one, and a shape inside a theme's style matrix can say
 * `phClr` ("the colour my caller passes me"). Narrowing to the twelve theme slots
 * would silently coerce two thirds of real references to the wrong token.
 */
export type SchemeToken = ThemeColorSlot | ColorMapToken | 'phClr'

export type Color =
	| { kind: 'srgb'; hex: string; alpha?: number }
	| {
			kind: 'scheme'
			slot: SchemeToken
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
 * A fill.
 *
 * `none` and `inherit` are separate states rather than one absent value, because
 * they are different decks: `none` is an explicit `a:noFill` (a deliberately
 * transparent shape), `inherit` is a shape that states no fill at all and takes
 * one from its `p:style/a:fillRef` or its placeholder chain. The legacy IR
 * overloaded `null` to mean both, so a themed shape and a transparent one were
 * indistinguishable.
 *
 * Import states both for a *shape*, off `Shape.fillNoFill` (ts-pptx 3.0.0,
 * closing {@link https://github.com/shbernal/ts-pptx/issues/1}), and for a *table
 * cell*, off `TableCell.fillNoFill` (3.1.0,
 * {@link https://github.com/shbernal/ts-pptx/issues/7}). Before that a suppressed
 * cell was indistinguishable from one inheriting the table style's shading, and
 * both landed on `inherit`.
 */
export type Fill =
	| { kind: 'inherit' }
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

/** An arrowhead (`a:headEnd` / `a:tailEnd`). `@w`/`@len` are optional in the schema. */
export interface LineEnd {
	type: 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow'
	width?: 'sm' | 'med' | 'lg'
	length?: 'sm' | 'med' | 'lg'
}

/**
 * An outline: shape border, connector line, or table cell edge.
 *
 * Three states for the same reason {@link Fill} has them — `none` is an explicit
 * `a:noFill` on the line, `inherit` is a shape that states none and takes one
 * from the theme's `a:lnRef`. Here the read model *does* distinguish the two
 * (`Shape.lineNoFill`), so both are reachable from import.
 *
 * `a:ln/@cap` is modeled and `@algn` is not, and the split is the round-trip rule
 * rather than a reading one — ts-pptx 3.0.0 exposes both
 * ({@link https://github.com/shbernal/ts-pptx/issues/2}), but only `@cap` has a
 * write-API option (`ShapeLineProps.cap`) to come back through. A stated `@algn`
 * is imported as a {@link FidelityNote}, which is what keeps it a *declared*
 * difference rather than a field that silently never survives.
 */
export type Stroke =
	| { kind: 'inherit' }
	| { kind: 'none' }
	| {
			kind: 'line'
			/** `a:ln/@w`. Absent when the line states no width and inherits one. */
			widthPt?: number
			/** Absent when the line states no colour of its own. */
			color?: Color
			gradient?: Gradient
			/** `a:prstDash/@val`. Absent when unstated; a token outside {@link DashStyle} is a note. */
			dash?: DashStyle
			/**
			 * `a:ln/@cap`, as the raw OOXML token. Absent when unstated — PowerPoint
			 * draws `flat`, but an explicit `flat` and an absent one are different
			 * decks for the same reason everything else here is optional.
			 *
			 * On a thick dashed rule the cap decides whether each dash reads as a
			 * rectangle or a lozenge and extends every dash by the stroke width, so it
			 * is geometry, not polish. SVG's `stroke-linecap` is the exact equivalent
			 * (`flat`→`butt`, `rnd`→`round`, `sq`→`square`).
			 */
			cap?: 'flat' | 'rnd' | 'sq'
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

/**
 * A paragraph's bullet definition — the four mutually exclusive `a:pPr` bullet
 * children (`a:buNone` / `a:buChar` / `a:buAutoNum` / `a:buBlip`).
 *
 * Every arm but `none` carries the bullet's own `a:buFont` / `a:buSzPct` /
 * `a:buClr`, which style the *glyph* rather than the text and are therefore not
 * reachable from the run properties. `none` carries none of them because there is
 * no glyph to style.
 *
 * `a:buSzPts` (an absolute glyph size) is deliberately unmodeled: the write API's
 * `bullet.size` is a percentage of the run size, so an absolute one has no way
 * back and is imported as a `text.bullet.sizePt` note instead.
 */
export type Bullet =
	| { kind: 'none' }
	| { kind: 'character'; char: string; font?: string; color?: Color; sizePct?: number }
	| {
			kind: 'number'
			/** `a:buAutoNum/@type`, e.g. `arabicPeriod`. */
			scheme: string
			/**
			 * `@startAt`. Absent means the schema default, 1. Numbering is content, not
			 * styling — a list continuing "5. Deploy" that comes back as "1. Deploy" is
			 * a different slide.
			 */
			startAt?: number
			font?: string
			color?: Color
			sizePct?: number
	  }
	| {
			/**
			 * `a:buBlip`, an image used as the glyph. Addressed by {@link AssetRef} like
			 * every other picture in this model, never re-embedded.
			 *
			 * Modeled rather than dropped because the read model resolves the image part
			 * and a renderer can paint it. It does not survive the emit leg — upstream's
			 * text mapper has no asset resolver and declares a `text.bullet.picture`
			 * note — so the loss is declared there rather than hidden here by pretending
			 * the paragraph inherited its bullet.
			 */
			kind: 'picture'
			asset: AssetRef
			font?: string
			color?: Color
			sizePct?: number
	  }

/**
 * Paragraph formatting.
 *
 * `align` and `bullet` are optional for the same reason every {@link RunProperties}
 * field is: absent means *inherited from the list style*, which is a different
 * paragraph from one that explicitly says `left` or `a:buNone`. `level` is not
 * optional — `a:pPr/@lvl` genuinely defaults to `0`.
 */
export interface ParagraphProperties {
	/** `@algn`, restricted to the four the write API expresses; `dist`/`thaiDist` are a note. */
	align?: 'left' | 'center' | 'right' | 'justify'
	/** `a:pPr/@lvl`, 0-based outline depth. */
	level: number
	bullet?: Bullet
	/** `@marL` / `@indent`, points. */
	marginLeftPt?: number
	indentPt?: number
	/** `a:lnSpc`, in whichever of the two OOXML forms the source used. */
	lineSpacing?: LineSpacing
	spaceBeforePt?: number
	spaceAfterPt?: number
}

/**
 * What a run's character formatting resolves to once the placeholder → layout →
 * master → `p:defaultTextStyle` chain has been walked.
 *
 * The counterpart to {@link RunProperties}, and the reason "absent means
 * inherited" is affordable: a renderer cannot paint an inherited value, so
 * without this the only way to draw a placeholder title at its real 44pt would
 * be to write 44 into the run — which is the flattening trap, and would bake a
 * layout's size into every slide the moment emit read it back.
 *
 * Only these four exist because these are the four the read model resolves.
 * Nothing may emit from them; they are paint, and {@link RunProperties} is what
 * the deck said.
 */
export interface ResolvedRunProperties {
	fontFace?: string
	sizePt?: number
	bold?: boolean
	color?: Color
}

export interface TextRun {
	text: string
	/** What this run itself stated. Absent field = inherited. */
	props: RunProperties
	/** What to paint. Derived at import; never written back to the deck. */
	resolved: ResolvedRunProperties
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
	/**
	 * `null` when nothing in the slide → layout → master chain gives this node a
	 * resolvable box — a non-placeholder shape with no `a:xfrm`, or a group whose
	 * `a:chExt` is degenerate. Rare, and honest: a renderer skips it and says so,
	 * which beats painting it at the origin as if that were stated.
	 */
	placement: Placement | null
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
	stroke: Stroke
	text: TextBody | null
}

export interface PictureNode extends NodeBase {
	kind: 'picture'
	asset: AssetRef
	/** `a:srcRect` crop. */
	crop?: EdgeRect
	geometry: Geometry
	stroke: Stroke
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
	stroke: Stroke
	start?: Connection
	end?: Connection
}

export interface TableColumn {
	/** `a:gridCol/@w`. `null` when the grid states none and the renderer distributes. */
	widthEmu: number | null
}

export interface TableCell {
	id: NodeId
	text: TextBody | null
	fill: Fill
	borders: {
		left: Stroke
		right: Stroke
		top: Stroke
		bottom: Stroke
	}
	/** `>1` when this cell spans; `null` for a plain 1×1 cell. */
	span: { columns: number; rows: number } | null
	/** `true` for a cell covered by a neighbour's span (`hMerge`/`vMerge`). */
	covered: boolean
	/**
	 * `a:tcPr/@marL`/`@marR`/`@marT`/`@marB`. EMU, not points, because that is the
	 * unit OOXML states them in — unlike {@link TextBody.insetsPt}, whose source
	 * attributes the read model already reports in points. Nothing here converts a
	 * unit; the field name says which one arrived.
	 *
	 * A side is `null` when the cell states none and inherits it from the style.
	 */
	marginsEmu: { left: number | null; right: number | null; top: number | null; bottom: number | null }
	anchor: 'top' | 'middle' | 'bottom'
}

export interface TableRow {
	/** `a:tr/@h`. `null` when the row height is left to the content. */
	heightEmu: number | null
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
 *
 * There is no child-space transform here on purpose. Every child's
 * {@link Placement} is already slide-absolute (see {@link Placement}), so a
 * renderer walks the tree for paint order and identity and never composes a
 * transform. Carrying `a:chOff`/`a:chExt` as well would be a second copy of the
 * same geometry that nothing keeps in step.
 *
 * That reason stands on its own now that the accessor exists: ts-pptx 3.0.0 added
 * `GroupShape.childFrame` ({@link https://github.com/shbernal/ts-pptx/issues/4}),
 * and it stays unread here deliberately. The ask was filed for a consumer that
 * *rebuilds* a group as OOXML and needs the source child space to reproduce its
 * scaling; this model paints, and `absoluteFrame` has already done that
 * arithmetic.
 */
export interface GroupNode extends NodeBase {
	kind: 'group'
	children: RenderNode[]
}

/**
 * A node this model cannot draw — a chart, a SmartArt diagram, an embedded
 * object. It exists so the renderer has something to put on the page, and it
 * always renders as a labelled placeholder.
 *
 * **Opaque is not the same as carried, and the two are independent.** Opaque
 * means *this IR cannot paint it*; {@link RenderSlide.source} `carried` means
 * *the write API cannot author it*. A plain chart is opaque and its slide is
 * authored — it round-trips through `addChart` perfectly and simply cannot be
 * drawn in a browser. A `chartEx` is both. Collapsing them would either carry
 * every slide holding a chart (losing an editable emitted deck for no reason) or
 * claim a `chartEx` survives transcription (which it does not).
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

/**
 * A slide's effective background.
 *
 * The fill is a plain {@link Fill} rather than a background-only union: a
 * background's fill options *are* a shape's, and a second near-identical type
 * would need its own mapper, its own renderer branch and its own reason to drift.
 * A `p:bgRef` (a background naming a style-matrix entry) resolves through the
 * theme to one of these before it gets here.
 */
export interface Background {
	/** Which tier of slide → layout → master the background resolved from. */
	source: 'slide' | 'layout' | 'master'
	fill: Fill
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
	/**
	 * Media the XML references. Keyed by relationship id, because the XML refers
	 * to its media by `r:embed`/`r:id` and nothing else — a bare list of assets
	 * would carry the bytes and lose the only thing that says which `<a:blip>`
	 * wants them. This is the shape `ExtractedSlide.media` takes on the way back.
	 */
	assets: ResidualAsset[]
}

export interface ResidualAsset {
	/** The `r:id` used inside {@link Residual.xml}. */
	relId: string
	asset: AssetRef
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
