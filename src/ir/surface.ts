/**
 * The **editable surface**: the exact set of things a human may change in the
 * rendered HTML and have honoured on the way back.
 *
 * Invariant R is scoped to the unedited loop *plus edits made through this
 * surface*. An edit outside it is **detected as drift**, never interpreted — the
 * slide drops to the heuristic lane and says so. That is only meaningful if
 * "outside" has a definition, which is what this file is.
 *
 * ## Why it is data and not prose
 *
 * The renderer writes the surface into the document (those are the nodes it
 * makes `contenteditable`) and the parser hashes the complement of it to detect
 * drift. If each side carried its own idea of what is editable, the two would
 * diverge and the failure mode would be silent: an edit the renderer invited but
 * the parser treats as drift, or worse, the reverse. So the list below is the single
 * source, and both {@link project} and {@link freeze} are derived from it rather
 * than restating it.
 *
 * ## v1 surface
 *
 * The smallest set that is useful: run text, the character properties that map
 * 1:1 onto a write-API option, the *paragraph* properties that do, and
 * deleting a node. Everything else — moving a box, changing geometry, restyling
 * a table, reordering or inserting slides — is out. Slide reordering is out for
 * a second reason as well: the writer has `removeSlide(index)` but no
 * `insertSlide`/`moveSlide`, so emit builds decks in append order.
 */

import { cloneIr, type NodeId, type NodeKind, type ParagraphProperties, type RenderIr, type RenderNode, type RunProperties } from './render'

/**
 * The character properties inside the surface, and the single place they are
 * named. {@link project} copies exactly these keys and {@link freeze} strips
 * exactly these keys, so adding one here is *nearly* the whole change — the
 * renderer, the reader and the edit mapper all derive from this list, but
 * `parse/edits.ts` still has to name the write-API spelling of each value and
 * `parse/surface.ts` still has to name each value's shape. Both are tested
 * against this list rather than against a second copy of it.
 *
 * They are the properties with a 1:1 write-API option. Anything needing
 * interpretation to get back into the deck — a font face that may not exist on
 * the target machine, a highlight, a baseline shift — is out of surface on
 * purpose: the return path must never have to guess.
 *
 * `underline` and `strike` qualify because {@link RunProperties} already models
 * each as the three-value subset the writer can express, and every other
 * `ST_TextUnderlineType` / `ST_TextStrikeType` token is imported as a fidelity
 * note rather than rounded into one of the three. Both include the *explicit*
 * off — `u="none"`, `strike="noStrike"` — which is a different fact from saying
 * nothing, and the reason `parse/edits.ts` maps them by table rather than
 * passing the value through.
 */
export const EDITABLE_RUN_PROPS = ['bold', 'italic', 'underline', 'strike', 'sizePt', 'color'] as const

export type EditableRunProp = (typeof EDITABLE_RUN_PROPS)[number]

/**
 * The paragraph properties inside the surface — the same 1:1 bar, one tier up.
 *
 * `align` clears it exactly. {@link ParagraphProperties.align} is the four values
 * `TextBaseProps.align` declares and no others (`dist`/`thaiDist` are imported as
 * a `text.align` note rather than rounded into one of the four), and the writer
 * omits `a:pPr/@algn` entirely when the option is absent — so *inherited*, *left*
 * and *centre* are three states the option can say and the file can hold.
 *
 * ## `bullet`, and the state it was waiting for
 *
 * `bullet` was out of surface until ts-pptx#15 shipped, and why is worth keeping,
 * because it is the general form of the bar: **the 1:1 test is per *value*, not
 * per property.** {@link ParagraphProperties.bullet} models *inherited* as
 * absence, `{kind:'none'}` as the explicit `a:buNone`, and a glyph as itself. The
 * write API could spell the second and the third and not the first — an omitted
 * `bullet` and `bullet: false` emitted byte-identical
 * `<a:pPr indent="0" marL="0"><a:buNone/></a:pPr>` — so a control offering
 * *inherited* would have written the explicit off, and nothing would have looked
 * wrong: a suppressed bullet and an inherited-none paint the same, so the lie
 * would have surfaced only later, when someone edited the master and the slide
 * stopped following it. `bullet: 'inherit'` states nothing at all, and the three
 * states are three spellings.
 *
 * ## What may be set is narrower than what is carried
 *
 * The projection carries whatever the paragraph holds, including the two glyphs
 * the write API cannot author back — a numbering scheme outside its sixteen, a
 * picture bullet, a glyph with its own theme colour. That asymmetry is safe
 * because **only the delta is ever written**: a bullet nobody touched comes back
 * through the projection unchanged and is therefore never re-authored, and the
 * `DeckIr` keeps the spelling `readModelToIr` gave it. `parse/edits.ts` draws the
 * line, and warns rather than approximating when a caller crosses it.
 *
 * ## The two margins, which `bullet` was hiding
 *
 * `marginLeftPt` and `indentPt` are `a:pPr/@marL` and `@indent` — where the body
 * text starts, and how far the first line is offset from it. They clear the bar
 * for the same reason `bullet` now does and could not before it: `paraMarginLeft`
 * and `paraIndent` take a number *or* `'inherit'`, so *36pt*, *explicitly zero*
 * and *whatever the list style says* are three states the option can spell.
 *
 * Omitting the option is none of those three. It means *the bullet's default* — a
 * drawn glyph writes its own hanging margin, `bullet: false` writes `marL="0"
 * indent="0"`, and only `bullet: 'inherit'` writes nothing — which is why the
 * margins could not be in surface while the bullet decided them. They are separate
 * facts now: an explicitly bulletless paragraph can keep the margin it inherits,
 * and a bulleted one can state its own.
 *
 * The value domain is a point measurement rather than an enumeration, so the
 * per-value bar lands on the range instead: `@marL` is unsigned and both attributes
 * stop at 4032pt, and the writer *clamps* a value outside that rather than
 * refusing it. A clamped value is an approximation the surface does not make, so
 * `parse/edits.ts` refuses one — the same warning-and-keep the bullets get.
 */
export const EDITABLE_PARA_PROPS = ['align', 'bullet', 'marginLeftPt', 'indentPt'] as const

export type EditableParaProp = (typeof EDITABLE_PARA_PROPS)[number]

/** One rule of the surface. Paths are dotted, relative to a `RenderSlide`. */
export interface SurfaceRule {
	kind: 'text' | 'runProp' | 'paraProp' | 'delete'
	path: string
	why: string
}

const PARAGRAPHS = 'nodes[].text.paragraphs[]'
const RUNS = `${PARAGRAPHS}.runs[]`

/**
 * The surface, in full. Table cell text and text inside groups are reached by
 * the same paths — {@link project} descends into `rows[].cells[].text` and
 * `children[]`, because a run is a run wherever it lives.
 */
export const EDITABLE_SURFACE: readonly SurfaceRule[] = [
	{
		kind: 'text',
		path: `${RUNS}.text`,
		why: 'The point of the exercise: changing the words. A run is the largest span with uniform formatting, so replacing its string cannot disturb anything else.',
	},
	...EDITABLE_RUN_PROPS.map(
		(prop): SurfaceRule => ({
			kind: 'runProp',
			path: `${RUNS}.props.${prop}`,
			why: 'Maps 1:1 onto a write-API option, so the return path sets it without interpreting anything.',
		})
	),
	...EDITABLE_PARA_PROPS.map(
		(prop): SurfaceRule => ({
			kind: 'paraProp',
			path: `${PARAGRAPHS}.props.${prop}`,
			why: 'Maps 1:1 onto a write-API option, and unlike the rest of `a:pPr` the option can also state nothing, which is what an absent key means.',
		})
	),
	{
		kind: 'delete',
		path: 'nodes[]',
		why: 'Removing a node needs no interpretation — the emitted deck simply makes one fewer call. Adding or moving one does, so neither is in surface.',
	},
]

// ---------------------------------------------------------------------------
// The sanctioned projection
// ---------------------------------------------------------------------------

/** A run, reduced to the parts of it that may change. */
export interface ProjectedRun {
	/** Address within the slide: which node, which paragraph, which run. */
	node: NodeId
	paragraph: number
	run: number
	text: string
	/** Only the {@link EDITABLE_RUN_PROPS}, and only those the run actually states. */
	props: Pick<RunProperties, EditableRunProp>
}

/**
 * A paragraph, reduced the same way.
 *
 * Present for *every* paragraph of a text-bearing node, including one with no
 * runs. A blank line is a paragraph in the source and can state an alignment
 * like any other, so a projection keyed off runs alone would have made exactly
 * the paragraphs with nothing in them uneditable.
 */
export interface ProjectedParagraph {
	node: NodeId
	paragraph: number
	/** Only the {@link EDITABLE_PARA_PROPS}, and only those the paragraph states. */
	props: Pick<ParagraphProperties, EditableParaProp>
}

export interface ProjectedNode {
	id: NodeId
	kind: NodeKind
	paragraphs: ProjectedParagraph[]
	runs: ProjectedRun[]
}

export interface ProjectedSlide {
	number: number
	/** In document order, so the projection is stable for hashing. */
	nodes: ProjectedNode[]
}

export interface SanctionedProjection {
	irVersion: number
	slides: ProjectedSlide[]
}

/**
 * Reduce an IR to what the surface permits: every editable value, addressed by
 * node id rather than by position.
 *
 * This is what the renderer marks as editable regions and what the parser reads
 * back.
 * A difference between the projection that went out and the one that came back
 * is an *edit*; a difference anywhere else is *drift* — see {@link freeze}.
 *
 * `RenderSlide.chrome` is **not** walked. Those shapes belong to the layout and
 * master parts, which the emitted deck binds to rather than redraws, so an edit
 * to one has nowhere to go — and offering it would be the surface promising
 * something the return path cannot honour. {@link freeze} keeps them, since a
 * change to them is drift.
 */
export function project(ir: RenderIr): SanctionedProjection {
	return {
		irVersion: ir.irVersion,
		slides: ir.slides.map((slide) => ({
			number: slide.number,
			nodes: slide.nodes.flatMap(projectNode),
		})),
	}
}

function projectNode(node: RenderNode): ProjectedNode[] {
	// A group contributes nothing of its own but everything its children hold;
	// flattening here is what keeps addressing by node id rather than by path.
	if (node.kind === 'group') return node.children.flatMap(projectNode)

	const paragraphs: ProjectedParagraph[] = []
	const runs: ProjectedRun[] = []
	if (node.kind === 'shape') collectText(node.id, node.text, paragraphs, runs)
	if (node.kind === 'table') {
		for (const row of node.rows) {
			for (const cell of row.cells) collectText(cell.id, cell.text, paragraphs, runs)
		}
	}
	return [{ id: node.id, kind: node.kind, paragraphs, runs }]
}

/** The shape {@link collectText} needs; both `ShapeNode.text` and `TableCell.text` satisfy it. */
type TextBodyLike = {
	paragraphs: { props: ParagraphProperties; runs: { text: string; props: RunProperties }[] }[]
} | null

function collectText(
	owner: NodeId,
	text: TextBodyLike,
	intoParagraphs: ProjectedParagraph[],
	intoRuns: ProjectedRun[]
): void {
	if (!text) return
	text.paragraphs.forEach((paragraph, paragraphIndex) => {
		intoParagraphs.push({
			node: owner,
			paragraph: paragraphIndex,
			props: editableParaProps(paragraph.props),
		})
		paragraph.runs.forEach((run, runIndex) => {
			intoRuns.push({
				node: owner,
				paragraph: paragraphIndex,
				run: runIndex,
				text: run.text,
				props: editableRunProps(run.props),
			})
		})
	})
}

/**
 * A run's stated properties, narrowed to the surface.
 *
 * Exported because the renderer writes exactly this onto each run's span
 * (`data-pxh-props`) and the return path reads exactly this back. That attribute
 * is not a duplicate of the island: the *painted* style on a run is
 * `props.X ?? resolved.X`, so reading a colour back off the rendered span would
 * promote an inherited value into a stated one on every placeholder — the
 * flattening this model is built to avoid. What the deck stated is a different
 * fact from what the browser painted, and only the first is in surface.
 */
export function editableRunProps(props: RunProperties): Pick<RunProperties, EditableRunProp> {
	const picked: Pick<RunProperties, EditableRunProp> = {}
	for (const key of EDITABLE_RUN_PROPS) {
		// Assign only what is stated. An absent key means *inherited*, and writing
		// `undefined` instead would both break the JSON island's "absence has one
		// spelling" rule and turn inheritance into an explicit value.
		if (props[key] !== undefined) Object.assign(picked, { [key]: props[key] })
	}
	return picked
}

/**
 * A paragraph's stated properties, narrowed to the surface.
 *
 * The counterpart of {@link editableRunProps}, and exported for the same reason:
 * the renderer writes exactly this onto each `<p>` (`data-pxh-paraprops`) and the
 * return path reads exactly this back. `level` stays out because `@lvl` is an index
 * into a list style the paint model does not carry, and `spaceBeforePt` /
 * `spaceAfterPt` because the option treats `0` as unset — a paragraph that states
 * zero space to suppress its list style's cannot be authored back, which is
 * upstream's own `text.paragraphSpaceZero` note and the per-value bar again.
 */
export function editableParaProps(
	props: Pick<ParagraphProperties, EditableParaProp>
): Pick<ParagraphProperties, EditableParaProp> {
	const picked: Pick<ParagraphProperties, EditableParaProp> = {}
	for (const key of EDITABLE_PARA_PROPS) {
		if (props[key] !== undefined) Object.assign(picked, { [key]: props[key] })
	}
	return picked
}

// ---------------------------------------------------------------------------
// The complement
// ---------------------------------------------------------------------------

/**
 * Value equality for two surface values.
 *
 * The one decision behind "did this document change", and it is made once
 * because both callers make it about the same values: `reconcile` uses it to
 * decide whether a slide is `exact` or `reconciled`, and `editsBetween` uses it
 * to decide whether a property becomes an edit at all. Two spellings of it would
 * be two answers to that question.
 *
 * `JSON.stringify` is exact here rather than approximate: every surface value is
 * a boolean, a number, a string, a {@link Color} or a {@link Bullet}, all of
 * which are JSON-safe by the IR's own contract, and both sides are produced by
 * the same projection with the same key order. `===` alone would report a bullet
 * nobody touched as edited on every round trip, because both sides arrive through
 * a clone or a JSON parse.
 *
 * `undefined` and `null` are normalised together, inside rather than at each call
 * site: an absent property and one holding `null` are the same statement here,
 * and the callers reached that conclusion separately before this did.
 */
export function sameSurfaceValue(a: unknown, b: unknown): boolean {
	return a === b || JSON.stringify(a ?? null) === JSON.stringify(b ?? null)
}

/**
 * The IR with every editable value stripped — the half that is *not* allowed to
 * change.
 *
 * The parser hashes this. Hashing the whole IR would flag a legitimate text edit as
 * drift; hashing only the projection would catch nothing, since the projection
 * is exactly the part that is meant to move. Deriving both from
 * {@link EDITABLE_SURFACE} is what keeps the two halves complementary by
 * construction rather than by review.
 *
 * Node *deletion* is in surface, so a frozen slide's node list is compared as a
 * subsequence, not for equality — this function does not decide that; it only
 * removes the values a deletion-aware comparison must not look at.
 */
export function freeze(ir: RenderIr): RenderIr {
	const clone = cloneIr(ir)
	for (const slide of clone.slides) {
		for (const node of slide.nodes) stripNode(node)
	}
	return clone
}

function stripNode(node: RenderNode): void {
	if (node.kind === 'group') {
		for (const child of node.children) stripNode(child)
		return
	}
	if (node.kind === 'shape') stripText(node.text)
	if (node.kind === 'table') {
		for (const row of node.rows) {
			for (const cell of row.cells) stripText(cell.text)
		}
	}
}

function stripText(text: TextBodyLike): void {
	if (!text) return
	for (const paragraph of text.paragraphs) {
		for (const key of EDITABLE_PARA_PROPS) delete paragraph.props[key]
		for (const run of paragraph.runs) {
			run.text = ''
			for (const key of EDITABLE_RUN_PROPS) delete run.props[key]
		}
	}
}
