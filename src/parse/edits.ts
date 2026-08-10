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
import type { Color, NodeId, RenderIr, RenderNode, RunProperties, TextBody } from '../ir/render'
import { EDITABLE_RUN_PROPS, type EditableRunProp } from '../ir/surface'

/** The write-API option each surface property is spelled as. */
const OPTION_OF: Record<EditableRunProp, string> = {
	bold: 'bold',
	italic: 'italic',
	underline: 'underline',
	strike: 'strike',
	sizePt: 'fontSize',
	color: 'color',
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

export interface EditSet {
	/** `nodeId/paragraph/run` → what changed. */
	runs: Map<string, RunEdit>
	deleted: Set<NodeId>
}

export function isEmpty(edits: EditSet): boolean {
	return edits.runs.size === 0 && edits.deleted.size === 0
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
		diffNodes(slide.nodes, indexById(now.nodes), surviving, runs, deleted)
	}

	return { runs, deleted }
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
	runs: Map<string, RunEdit>,
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
			diffNodes(node.children, after, surviving, runs, deleted)
			continue
		}
		if (node.kind === 'shape' && now.kind === 'shape') diffText(node.id, node.text, now.text, runs)
		if (node.kind === 'table' && now.kind === 'table') {
			node.rows.forEach((row, rowIndex) => {
				row.cells.forEach((cell, cellIndex) => {
					const cellNow = now.rows[rowIndex]?.cells[cellIndex]
					if (cellNow !== undefined) diffText(cell.id, cell.text, cellNow.text, runs)
				})
			})
		}
	}
}

function diffText(owner: NodeId, before: TextBody | null, after: TextBody | null, runs: Map<string, RunEdit>): void {
	if (before === null || after === null) return
	before.paragraphs.forEach((paragraph, paragraphIndex) => {
		paragraph.runs.forEach((run, runIndex) => {
			const now = after.paragraphs[paragraphIndex]?.runs[runIndex]
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
				runs.set(`${owner}/${paragraphIndex}/${runIndex}`, edit)
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
 * The one place a run's new value is actually written.
 *
 * The paint model nests runs in paragraphs and the write API states them flat,
 * with the paragraph break carried as `breakLine` on the run *before* it. So the
 * join is document order, and it is guarded by a count: if the two disagree about
 * how many runs a frame has, no edit is placed at all.
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
	text.paragraphs.forEach((paragraph, paragraphIndex) => {
		paragraph.runs.forEach((_run, runIndex) => {
			addresses.push(`${owner}/${paragraphIndex}/${runIndex}`)
		})
	})
	if (!addresses.some((address) => edits.runs.has(address))) return

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
		if (edit === undefined) return
		const entry = runArgs[index] as { text?: unknown; options?: Record<string, unknown> }
		if (edit.text !== undefined) entry.text = edit.text
		if (Object.keys(edit.props).length === 0) return

		const options = entry.options ?? {}
		entry.options = options
		for (const [key, value] of Object.entries(edit.props)) {
			const option = OPTION_OF[key as EditableRunProp]
			if (value === undefined) delete options[option]
			else options[option] = value
		}
	})
}
