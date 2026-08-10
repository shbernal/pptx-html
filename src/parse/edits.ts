/**
 * Carrying a document's edits from the paint model into the write contract.
 *
 * The two IRs never meet — `src/emit/` may not import `src/ir/render.ts`, so the
 * renderer can never influence what is written — and this file is the one seam
 * where the boundary is crossed, in one direction, carrying one thing: **the
 * difference an edit made.**
 *
 * ## Only the delta is ever applied
 *
 * Nothing here writes a `RenderIr` value into a `DeckIr` unless that value
 * *changed*. That is the property the whole file is built around, and it is
 * stronger than it looks:
 *
 * - An unedited document exercises none of the structural mapping below, so
 *   Invariant R for the unedited loop does not depend on a single assumption in
 *   this file. It reduces to `import → emit`, which the oracle already gates.
 * - The two models disagree about defaults on purpose — `readModelToIr` resolves
 *   a run's colour to `000000` where the paint model correctly records that the
 *   run stated nothing. Copying wholesale would flatten every one of those into
 *   an explicit value; copying only what moved cannot.
 *
 * ## The alignment, and why it is checked rather than trusted
 *
 * `DeckIr.slides[i].calls[j]` and `RenderIr.slides[i].nodes[j]` line up because
 * both are built by walking the same `slide.shapes` in document order. That is
 * true of upstream's mapper and of this package's, but it is a property of two
 * independent traversals rather than a contract either one publishes — and
 * `CallIr` carries no `p:cNvPr/@id` to join on, only `sourceName`.
 *
 * So the alignment is *asserted*: same length, and every call's `sourceName`
 * matching its node's name. A mismatch throws. The alternative — patching the
 * run at the index that lined up — would put a user's text into some other
 * shape, which is the worst failure this whole architecture is arranged to
 * prevent, and it would look exactly like success.
 */

import type { CallIr, DeckIr, IrValue } from '@shbernal/ts-pptx/script'
import type {
	Bullet,
	Color,
	NodeId,
	ParagraphProperties,
	RenderIr,
	RenderNode,
	RunProperties,
	TextBody,
} from '../ir/render'
import { EDITABLE_PARA_PROPS, EDITABLE_RUN_PROPS, type EditableParaProp, type EditableRunProp } from '../ir/surface'

/** The write-API option each surface property is spelled as. */
const OPTION_OF: Record<EditableRunProp, string> = {
	bold: 'bold',
	italic: 'italic',
	underline: 'underline',
	strike: 'strike',
	sizePt: 'fontSize',
	color: 'color',
}

/** The same, for the paragraph tier. */
const PARA_OPTION_OF: Record<EditableParaProp, string> = {
	align: 'align',
	bullet: 'bullet',
	marginLeftPt: 'paraMarginLeft',
	indentPt: 'paraIndent',
}

/**
 * How each paragraph property says **nothing** — the value to write when an edit
 * returns one to inherited.
 *
 * This table exists because the answers differ, and the difference is the whole of
 * ts-pptx#15. Omitting `align` omits `a:pPr/@algn`, so *absent* and *inherited* are
 * the same act and the option is simply deleted. Omitting `bullet` does not omit
 * anything: it means `false`, which writes an explicit `<a:buNone/>` plus
 * `indent="0" marL="0"` and *overrides* the list style — silently, since a
 * suppressed bullet and an inherited-none paint identically. `'inherit'` is the
 * spelling for saying nothing, and it is what `readModelToIr` itself emits for such
 * a paragraph, so an edit that clears a bullet leaves the contract in the shape a
 * fresh read would have produced.
 *
 * The two margins take `'inherit'` for the same reason, and it is the same
 * sentence one attribute along: an omitted `paraMarginLeft` is not silence either,
 * it is *the bullet's default* — the hanging margin of a drawn glyph, or the
 * `marL="0"` that `bullet: false` writes. Deleting the option on a paragraph whose
 * bullet is anything but inherited would move its text, which is the one failure
 * the surface cannot afford: an edit made through it, accepted, and quietly turned
 * into something else.
 */
const PARA_INHERITED_OF: Record<EditableParaProp, unknown> = {
	align: undefined,
	bullet: 'inherit',
	marginLeftPt: 'inherit',
	indentPt: 'inherit',
}

/**
 * Which runs of a paragraph a property is written to — and two of the answers are
 * opposites, which is why this is a table and not a rule.
 *
 * The write API has no paragraph tier: properties ride on runs, and
 * `groupRunsIntoLines` decides where one paragraph ends and the next begins by
 * looking at them. It uses two different signals, and each property has to be
 * placed for the signal that reads it:
 *
 * - **`align` on every run.** Two adjacent runs that disagree about their
 *   alignment start a new paragraph, so setting it on the first run of a three-run
 *   paragraph does not restyle that paragraph, it *splits* it.
 * - **`bullet` on the first run only.** A run that states a *glyph* starts a new
 *   paragraph on its own, disagreement or not. So the placement that is right for
 *   `align` is exactly wrong here: writing one glyph to all three runs produces
 *   three one-run paragraphs. `false` and `'inherit'` are inert either way, but
 *   the first run is the only placement correct for all three states.
 * - **The two margins on every run**, and this one is a choice rather than a
 *   constraint. The grouper does not look at them at all and the serializer reads
 *   `a:pPr` off whichever run opens the line, so either placement emits the same
 *   attribute. Every run is what `readModelToIr` produces, and matching it is worth
 *   more than the shorter write: it keeps one shape of contract for a paragraph
 *   rather than one the edit path recognises and one a fresh read does.
 *
 * All three are that shape — upstream replicates a paragraph's options onto every
 * run and deletes `bullet` from the continuations — so the contract this leaves
 * behind is the one a fresh read of the same deck would have written.
 *
 * None of these mistakes shows up in a run count, which is what the guard in
 * {@link patchRuns} checks. What catches them is asserting the *paragraph* count
 * after the round trip, which is what the corpus decks with two runs in one
 * paragraph exist for.
 */
const PARA_PLACEMENT_OF: Record<EditableParaProp, 'every-run' | 'first-run'> = {
	align: 'every-run',
	bullet: 'first-run',
	marginLeftPt: 'every-run',
	indentPt: 'every-run',
}

/**
 * `RunProperties.underline` in the write API's spelling. `TextPropsOptions.underline`
 * is an object so it can carry `a:uFill` beside the style; the style alone is
 * what this surface offers, and the three tokens are exactly the three modeled.
 */
const UNDERLINE_OPTION: Record<NonNullable<RunProperties['underline']>, string> = {
	none: 'none',
	single: 'sng',
	double: 'dbl',
}

/**
 * `RunProperties.strike`, likewise — all three of them the option's own declared
 * values, since ts-pptx 3.1.0+b16fb74b widened the union to carry `noStrike`.
 *
 * The `none` arm is the one that has to be argued for, and the argument is the
 * same one the whole surface rests on. `TextPropsOptions.strike` also accepts
 * `false`, and `false` *omits* `a:rPr/@strike` — which states nothing and leaves
 * the run whatever it inherits. `noStrike` states off. Mapping `none` to `false`
 * would turn "this run is deliberately not struck through" into "this run says
 * nothing", the flattening the surface exists to refuse.
 *
 * Nothing on this path would fail to compile if the token stopped reaching the
 * attribute — a `CallIr` argument is an `IrValue`, so the option's type is not
 * checked here at all. The `text-decoration` corpus deck and the oracle's
 * `explicit "not underlined"` test are what hold it.
 */
const STRIKE_OPTION: Record<NonNullable<RunProperties['strike']>, string> = {
	none: 'noStrike',
	single: 'sngStrike',
	double: 'dblStrike',
}

/** One run's new state, addressed as the surface addresses it. */
interface RunEdit {
	text?: string
	props: Partial<Record<EditableRunProp, unknown>>
}

/** One paragraph's new state. `undefined` for a value means *back to inherited*. */
interface ParaEdit {
	props: Partial<Record<EditableParaProp, unknown>>
}

export interface EditSet {
	/** `nodeId/paragraph/run` → what changed. */
	runs: Map<string, RunEdit>
	/** `nodeId/paragraph` → what changed. */
	paragraphs: Map<string, ParaEdit>
	deleted: Set<NodeId>
}

export function isEmpty(edits: EditSet): boolean {
	return edits.runs.size === 0 && edits.paragraphs.size === 0 && edits.deleted.size === 0
}

/**
 * What changed between the model as rendered and the model as it came back.
 *
 * Computed over the two `RenderIr`s rather than read off the reconciliation,
 * because a caller may have edited the model directly — and because a delta
 * derived from the two things being compared cannot disagree with them.
 */
export function editsBetween(before: RenderIr, after: RenderIr): EditSet {
	const runs = new Map<string, RunEdit>()
	const paragraphs = new Map<string, ParaEdit>()
	const deleted = new Set<NodeId>()

	for (const slide of before.slides) {
		const now = after.slides.find((entry) => entry.number === slide.number)
		if (now === undefined) {
			// Slide removal is not in the editable surface; a caller that dropped one
			// has left the loop and is not owed a reconstruction of what it meant.
			throw new Error(`slide ${slide.number} is missing from the edited model, and slide removal is not in the editable surface`)
		}
		const surviving = new Set<NodeId>()
		collectIds(now.nodes, surviving)
		diffNodes(slide.nodes, indexById(now.nodes), surviving, { runs, paragraphs }, deleted)
	}

	return { runs, paragraphs, deleted }
}

/** The two address maps `diffNodes` fills, carried together so the walk stays readable. */
interface TextEdits {
	runs: Map<string, RunEdit>
	paragraphs: Map<string, ParaEdit>
}

function collectIds(nodes: readonly RenderNode[], into: Set<NodeId>): void {
	for (const node of nodes) {
		into.add(node.id)
		if (node.kind === 'group') collectIds(node.children, into)
	}
}

function indexById(nodes: readonly RenderNode[]): Map<NodeId, RenderNode> {
	const index = new Map<NodeId, RenderNode>()
	const walk = (list: readonly RenderNode[]): void => {
		for (const node of list) {
			index.set(node.id, node)
			if (node.kind === 'group') walk(node.children)
		}
	}
	walk(nodes)
	return index
}

function diffNodes(
	nodes: readonly RenderNode[],
	after: ReadonlyMap<NodeId, RenderNode>,
	surviving: ReadonlySet<NodeId>,
	edits: TextEdits,
	deleted: Set<NodeId>
): void {
	for (const node of nodes) {
		if (!surviving.has(node.id)) {
			deleted.add(node.id)
			continue
		}
		const now = after.get(node.id)
		if (now === undefined || now.kind !== node.kind) continue

		if (node.kind === 'group' && now.kind === 'group') {
			diffNodes(node.children, after, surviving, edits, deleted)
			continue
		}
		if (node.kind === 'shape' && now.kind === 'shape') diffText(node.id, node.text, now.text, edits)
		if (node.kind === 'table' && now.kind === 'table') {
			node.rows.forEach((row, rowIndex) => {
				row.cells.forEach((cell, cellIndex) => {
					const cellNow = now.rows[rowIndex]?.cells[cellIndex]
					if (cellNow !== undefined) diffText(cell.id, cell.text, cellNow.text, edits)
				})
			})
		}
	}
}

/**
 * Restate the margins when an edit stops the bullet being inherited.
 *
 * The one place a paragraph edit changes the meaning of an option it did not
 * write. An omitted `paraMarginLeft` is *the bullet's default*, so what the
 * contract's silence about the margins means depends on the bullet beside them:
 * under `bullet: 'inherit'` it is silence too, and `readModelToIr` omits the keys
 * for exactly that paragraph. Change that bullet to `false` or to a glyph and the
 * same silence starts meaning `marL="0" indent="0"` or a hanging pair — so
 * suppressing an inherited bullet would flatten the paragraph's inherited indent
 * along with it, invisibly, having been asked for one thing.
 *
 * So a bullet edit carries the margins the paragraph *states* — a number, or
 * `'inherit'` for one it does not state. That is what a fresh read of the result
 * would produce, since upstream's own mapper omits the keys only while the bullet
 * is inherited. Nothing is pinned when the margin was edited too (the edit already
 * says it), nor when the bullet edit is one the write API cannot author (nothing is
 * written, so nothing changes meaning).
 */
function pinMarginsBesideBullet(edit: ParaEdit, props: ParagraphProperties): void {
	const bullet = edit.props.bullet
	if (bullet === undefined || bullet === 'inherit' || whyUnspellable(bullet) !== null) return
	for (const key of ['marginLeftPt', 'indentPt'] as const) {
		if (key in edit.props) continue
		const stated = props[key]
		edit.props[key] = stated === undefined ? PARA_INHERITED_OF[key] : paraOptionValue(key, stated)
	}
}

function diffText(owner: NodeId, before: TextBody | null, after: TextBody | null, edits: TextEdits): void {
	if (before === null || after === null) return
	before.paragraphs.forEach((paragraph, paragraphIndex) => {
		const paragraphNow = after.paragraphs[paragraphIndex]
		if (paragraphNow !== undefined) {
			const paraEdit: ParaEdit = { props: {} }
			for (const key of EDITABLE_PARA_PROPS) {
				const wasSet = paragraph.props[key]
				const isSet = paragraphNow.props[key]
				// Structurally, like the run loop below. `bullet` is an object, and both
				// sides arrive through a clone or a JSON parse, so `===` would report a
				// bullet nobody touched as edited on every round trip.
				if (JSON.stringify(wasSet ?? null) === JSON.stringify(isSet ?? null)) continue
				paraEdit.props[key] = isSet === undefined ? PARA_INHERITED_OF[key] : paraOptionValue(key, isSet)
			}
			pinMarginsBesideBullet(paraEdit, paragraphNow.props)
			if (Object.keys(paraEdit.props).length > 0) edits.paragraphs.set(`${owner}/${paragraphIndex}`, paraEdit)
		}

		paragraph.runs.forEach((run, runIndex) => {
			const now = paragraphNow?.runs[runIndex]
			if (now === undefined) return

			const edit: RunEdit = { props: {} }
			if (now.text !== run.text) edit.text = now.text
			for (const key of EDITABLE_RUN_PROPS) {
				const wasSet = run.props[key]
				const isSet = now.props[key]
				if (JSON.stringify(wasSet ?? null) === JSON.stringify(isSet ?? null)) continue
				edit.props[key] = isSet === undefined ? undefined : optionValue(key, isSet)
			}
			if (edit.text !== undefined || Object.keys(edit.props).length > 0) {
				edits.runs.set(`${owner}/${paragraphIndex}/${runIndex}`, edit)
			}
		})
	})
}

/**
 * A surface value in the write API's spelling.
 *
 * `bold`, `italic` and `sizePt` pass straight through; the other three do not,
 * and each for its own reason.
 *
 * Colour takes a bare hex string or a theme-slot name, which is exactly the two
 * arms of {@link Color} and nothing else. A colour carrying transforms keeps its
 * *slot*, not its resolved hex — writing the resolved value would bake the source
 * theme into the deck and stop the shape tracking it, which is the flattening the
 * paint model exists to avoid.
 *
 * `underline` and `strike` are the OOXML tokens rather than this model's names:
 * the IR spells the three values as `none`/`single`/`double` because that reads,
 * and the deck spells them `none`/`sng`/`dbl` and `noStrike`/`sngStrike`/`dblStrike`.
 * The translation is a table lookup with no third case, which is what keeps it a
 * spelling change rather than an interpretation.
 */
function optionValue(key: EditableRunProp, value: RunProperties[EditableRunProp]): unknown {
	if (key === 'color') {
		const color = value as Color
		return color.kind === 'srgb' ? color.hex : color.slot
	}
	if (key === 'underline') return { style: UNDERLINE_OPTION[value as NonNullable<RunProperties['underline']>] }
	if (key === 'strike') return STRIKE_OPTION[value as NonNullable<RunProperties['strike']>]
	return value
}

/**
 * A marker for a value the write API has no spelling for.
 *
 * A symbol key rather than a sentinel string or `null`, because both of those are
 * things a real option could be — and the one thing this must never do is get
 * written to the deck as if it were a value. {@link applyEdits} reports it and
 * {@link patchRuns} skips it, so the paragraph keeps whatever `readModelToIr`
 * gave it.
 */
const UNSPELLABLE = Symbol('no write-API spelling')

function unspellable(why: string): Record<symbol, string> {
	return { [UNSPELLABLE]: why }
}

function whyUnspellable(value: unknown): string | null {
	if (typeof value !== 'object' || value === null || !(UNSPELLABLE in value)) return null
	return (value as Record<symbol, string>)[UNSPELLABLE] as string
}

/**
 * A paragraph value in the write API's spelling.
 *
 * `align` passes straight through: {@link ParagraphProperties.align} was defined
 * as the option's own domain rather than as OOXML's (`l`/`ctr`/`r`/`just` are the
 * *attribute's* spelling, and the importer translates them on the way in).
 *
 * `bullet` is where the surface's per-value bar becomes visible in code. Three of
 * {@link Bullet}'s states are the three the option can author, and they are the
 * three {@link PARA_INHERITED_OF} completes:
 *
 * | model | option | file |
 * |---|---|---|
 * | absent | `'inherit'` | no `a:pPr` at all, so the list style still reaches it |
 * | `{kind:'none'}` | `false` | `<a:pPr indent="0" marL="0"><a:buNone/></a:pPr>` |
 * | `{kind:'character', char}` | `{characterCode}` | `<a:buChar/>` at that code point |
 *
 * The rest are refused rather than approximated, and each for a reason the surface
 * already gives elsewhere. A numbering scheme is a free `a:buAutoNum/@type` string
 * in this model and `numberType` declares sixteen of `ST_TextAutonumberScheme`'s
 * forty, so a paragraph could hold one the option cannot name. A picture bullet
 * addresses an `AssetRef` this package never re-embeds — upstream's own text
 * mapper says the same with its `text.bullet.picture` note. A glyph with its own
 * font, size or colour would have to flatten a `scheme` colour to the hex it
 * happens to resolve to, which is the theme-baking the paint model exists to
 * avoid, since `bullet.color` takes a hex and nothing else.
 *
 * None of those is reachable without a caller deliberately writing one, because
 * only the delta is applied and a bullet nobody moved is not a delta.
 *
 * The two margins pass through as their number, with the same bar applied to a
 * range rather than to a set — see {@link marginOptionValue}.
 */
function paraOptionValue(key: EditableParaProp, value: NonNullable<ParagraphProperties[EditableParaProp]>): unknown {
	if (key === 'align') return value
	if (key === 'marginLeftPt' || key === 'indentPt') return marginOptionValue(key, value as number)

	const bullet = value as Bullet
	if (bullet.kind === 'none') return false
	if (bullet.kind === 'number') {
		return unspellable(`a ${JSON.stringify(bullet.scheme)} numbered bullet`)
	}
	if (bullet.kind === 'picture') return unspellable('a picture bullet')
	if (bullet.font !== undefined || bullet.color !== undefined || bullet.sizePct !== undefined) {
		return unspellable('a bullet glyph carrying its own font, size or colour')
	}
	const points = [...bullet.char]
	if (points.length !== 1) return unspellable(`the ${points.length}-character glyph ${JSON.stringify(bullet.char)}`)
	return { characterCode: (bullet.char.codePointAt(0) as number).toString(16).toUpperCase().padStart(4, '0') }
}

/**
 * `a:pPr/@marL` is `ST_TextMargin`, which is unsigned and stops at 4032pt;
 * `@indent` is `ST_TextIndent`, signed and stopping at the same magnitude. A
 * negative margin is not a narrower one, it is a file PowerPoint offers to repair.
 */
const MARGIN_BOUNDS: Record<'marginLeftPt' | 'indentPt', { noun: string; attribute: string; min: number; max: number }> =
	{
		marginLeftPt: { noun: 'a margin', attribute: 'a:pPr/@marL', min: 0, max: 4032 },
		indentPt: { noun: 'an indent', attribute: 'a:pPr/@indent', min: -4032, max: 4032 },
	}

/**
 * A margin in the write API's spelling, which is the number itself — as long as the
 * attribute can hold it.
 *
 * This is the per-value bar on a measurement, where "which values can the option
 * author" is a range rather than a list. The writer *clamps* a value outside the
 * schema's, and warns as it does: `paraMarginLeft: -10` writes `marL="0"`, which is
 * a paragraph the caller did not ask for and the model does not hold. So an
 * out-of-range value is refused here instead, the deck keeps the margin
 * `readModelToIr` gave it, and the caller is told which value was dropped — the
 * same treatment a numbering scheme gets, for the same reason.
 *
 * Unreachable from the document: `parse/surface.ts` accepts any finite number, but
 * a rendered margin is one this package imported from a conforming file, so only a
 * caller editing the model can put a value out of range.
 */
function marginOptionValue(key: 'marginLeftPt' | 'indentPt', value: number): unknown {
	const { noun, attribute, min, max } = MARGIN_BOUNDS[key]
	if (!Number.isFinite(value)) return unspellable(`${noun} of ${JSON.stringify(value)}, which is not a measurement`)
	if (value < min || value > max) {
		return unspellable(`${noun} of ${value}pt, outside the ${min}pt to ${max}pt ${attribute} can hold`)
	}
	return value
}

// ---------------------------------------------------------------------------
// Applying them
// ---------------------------------------------------------------------------

export interface AppliedEdits {
	deck: DeckIr
	warnings: string[]
}

/**
 * Fold an edit set into a `DeckIr`.
 *
 * The deck is cloned, not mutated: the caller's copy is the one that came from
 * `readModelToIr` and is what a diff would be taken against.
 */
export function applyEdits(deck: DeckIr, structure: RenderIr, edits: EditSet): AppliedEdits {
	if (isEmpty(edits)) return { deck, warnings: [] }

	const next = structuredClone(deck) as DeckIr
	const warnings: string[] = []

	// Reported once per paragraph, before the walk, rather than inside `patchRuns`
	// — a paragraph property is written to every one of its runs, so warning where
	// it is written would say the same thing three times for a three-run paragraph.
	for (const [address, edit] of edits.paragraphs) {
		for (const [key, value] of Object.entries(edit.props)) {
			const why = whyUnspellable(value)
			if (why !== null) {
				warnings.push(
					`paragraph ${address}: ${key} was set to ${why}, which the write API cannot author; the deck's own ${key} was left in place`
				)
			}
		}
	}

	for (const slide of next.slides) {
		const nodes = structure.slides.find((entry) => entry.number === slide.number)?.nodes
		if (nodes === undefined) continue
		slide.calls = patchLevel(nodes, slide.calls, edits, warnings, slide.number)
	}

	return { deck: next, warnings }
}

/**
 * One level of the shape tree against one level of the call list.
 *
 * Returns the surviving calls, so a deletion is expressed as a filter rather than
 * an index-shifting splice mid-walk.
 */
function patchLevel(
	nodes: readonly RenderNode[],
	calls: readonly CallIr[],
	edits: EditSet,
	warnings: string[],
	slideNumber: number
): CallIr[] {
	if (calls.length !== nodes.length) {
		throw new Error(
			`slide ${slideNumber}: the write contract has ${calls.length} call(s) and the paint model has ${nodes.length} node(s); the two traversals no longer agree and an edit cannot be placed safely`
		)
	}

	const kept: CallIr[] = []
	nodes.forEach((node, index) => {
		const call = calls[index] as CallIr
		if (call.sourceName !== undefined && call.sourceName !== node.name) {
			throw new Error(
				`slide ${slideNumber}: call ${index} is for ${JSON.stringify(call.sourceName)} and node ${index} is ${JSON.stringify(node.name)}; an edit would land on the wrong shape`
			)
		}
		if (edits.deleted.has(node.id)) return
		patchNode(node, call.method, call.args, edits, warnings)
		kept.push(call)
	})
	return kept
}

/** The run-entry list a node's text lives in, or `null` when it has none. */
function patchNode(node: RenderNode, method: string, args: IrValue[], edits: EditSet, warnings: string[]): void {
	if (node.kind === 'group') {
		const children = args[0]
		if (Array.isArray(children)) patchGroup(node.children, children, edits, warnings)
		return
	}

	if (node.kind === 'shape') {
		// `addText` carries its runs as the first argument; a shape that happens to
		// hold text is written as `addShape(type, {…, text})`.
		const runs = method === 'addText' ? args[0] : (args[1] as Record<string, unknown> | undefined)?.text
		patchRuns(node.id, node.text, runs, edits, warnings)
		return
	}

	if (node.kind === 'table') {
		const rows = args[0]
		if (!Array.isArray(rows)) return
		node.rows.forEach((row, rowIndex) => {
			const rowArgs = rows[rowIndex]
			if (!Array.isArray(rowArgs)) return
			row.cells.forEach((cell, cellIndex) => {
				const cellArgs = rowArgs[cellIndex] as { text?: unknown } | undefined
				if (cellArgs === undefined) return
				patchRuns(cell.id, cell.text, cellArgs.text, edits, warnings)
			})
		})
	}
}

/**
 * A group's children, against the write API's child entries.
 *
 * `addGroup` takes `{shape}` / `{text}` / `{image}` / `{group}` wrappers rather
 * than a flat call list, so the discriminator has to be unwrapped before the same
 * alignment check applies. A child whose wrapper is not one of these is left
 * alone and reported: refusing to guess costs the edit, and guessing costs the
 * shape.
 */
function patchGroup(nodes: readonly RenderNode[], children: IrValue[], edits: EditSet, warnings: string[]): void {
	if (children.length !== nodes.length) {
		warnings.push(
			`a group holds ${children.length} child entr(ies) in the write contract and ${nodes.length} in the paint model; its edits were not applied`
		)
		return
	}
	nodes.forEach((node, index) => {
		const entry = children[index] as Record<string, { options?: Record<string, unknown>; text?: unknown; children?: IrValue[] }>
		if (node.kind === 'group') {
			const nested = entry.group?.children
			if (Array.isArray(nested)) patchGroup(node.children, nested, edits, warnings)
			return
		}
		if (node.kind !== 'shape') return
		const wrapper = entry.text ?? entry.shape
		if (wrapper === undefined) {
			warnings.push(`group child ${node.id} has no text entry in the write contract; its edits were not applied`)
			return
		}
		patchRuns(node.id, node.text, (wrapper as { text?: unknown }).text, edits, warnings)
	})
}

/**
 * The one place a run's or a paragraph's new value is actually written.
 *
 * The paint model nests runs in paragraphs and the write API states them flat,
 * with the paragraph break carried as `breakLine` on the run *before* it. So the
 * join is document order, and it is guarded by a count: if the two disagree about
 * how many runs a frame has, no edit is placed at all.
 *
 * ## Which runs a paragraph property lands on
 *
 * Not one answer: {@link PARA_PLACEMENT_OF} holds it per property, because
 * `align` has to go on every run of its paragraph and `bullet` on the first alone,
 * and each placement splits the paragraph if used for the other. Both failures
 * leave the run count untouched, so the guard above cannot see either.
 */
function patchRuns(
	owner: NodeId,
	text: TextBody | null,
	runArgs: unknown,
	edits: EditSet,
	warnings: string[]
): void {
	if (text === null) return
	const addresses: string[] = []
	// Which paragraph each flat run belongs to, so a paragraph edit can reach all
	// of that paragraph's entries and none of its neighbour's.
	const paragraphOf: number[] = []
	text.paragraphs.forEach((paragraph, paragraphIndex) => {
		paragraph.runs.forEach((_run, runIndex) => {
			addresses.push(`${owner}/${paragraphIndex}/${runIndex}`)
			paragraphOf.push(paragraphIndex)
		})
	})
	const touched =
		addresses.some((address) => edits.runs.has(address)) ||
		paragraphOf.some((index) => edits.paragraphs.has(`${owner}/${index}`))
	if (!touched) return

	if (!Array.isArray(runArgs)) {
		warnings.push(`${owner}: the write contract states its text in a form this build cannot address; its edits were not applied`)
		return
	}
	if (runArgs.length !== addresses.length) {
		warnings.push(
			`${owner}: the write contract has ${runArgs.length} run(s) and the paint model has ${addresses.length}; its edits were not applied`
		)
		return
	}

	addresses.forEach((address, index) => {
		const edit = edits.runs.get(address)
		const paraEdit = edits.paragraphs.get(`${owner}/${paragraphOf[index] ?? -1}`)
		if (edit === undefined && paraEdit === undefined) return
		const entry = runArgs[index] as { text?: unknown; options?: Record<string, unknown> }
		if (edit?.text !== undefined) entry.text = edit.text

		const props = { ...edit?.props }
		const paraProps = paraEdit?.props ?? {}
		if (Object.keys(props).length === 0 && Object.keys(paraProps).length === 0) return

		const options = entry.options ?? {}
		entry.options = options
		for (const [key, value] of Object.entries(props)) {
			const option = OPTION_OF[key as EditableRunProp]
			if (value === undefined) delete options[option]
			else options[option] = value
		}
		const opensParagraph = paragraphOf[index] !== paragraphOf[index - 1]
		for (const [key, value] of Object.entries(paraProps)) {
			const prop = key as EditableParaProp
			// Already reported by `applyEdits`; leaving the option untouched is what
			// keeps the deck's own value rather than a rounded-off version of it.
			if (whyUnspellable(value) !== null) continue
			if (PARA_PLACEMENT_OF[prop] === 'first-run' && !opensParagraph) continue
			const option = PARA_OPTION_OF[prop]
			if (value === undefined) delete options[option]
			else options[option] = value
		}
	})
}
