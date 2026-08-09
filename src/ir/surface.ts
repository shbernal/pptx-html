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
 * Part 05 renders the surface into the document (those are the nodes it makes
 * `contenteditable`) and part 06 hashes the complement of it to detect drift. If
 * each part carried its own idea of what is editable, the two would diverge and
 * the failure mode would be silent: an edit the renderer invited but the parser
 * treats as drift, or worse, the reverse. So the list below is the single
 * source, and both {@link project} and {@link freeze} are derived from it rather
 * than restating it.
 *
 * ## v1 surface
 *
 * The smallest set that is useful: run text, the four character properties that
 * map 1:1 onto a write-API option, and deleting a node. Everything else —
 * moving a box, changing geometry, restyling a table, reordering or inserting
 * slides — is out. Slide reordering is out for a second reason as well: the
 * writer has `removeSlide(index)` but no `insertSlide`/`moveSlide`, so emit
 * builds decks in append order.
 */

import type { NodeId, NodeKind, RenderIr, RenderNode, RunProperties } from './render'

/**
 * The character properties inside the surface, and the single place they are
 * named. {@link project} copies exactly these keys and {@link freeze} strips
 * exactly these keys, so adding one here is the whole change.
 *
 * They are the properties with a 1:1 write-API option. Anything needing
 * interpretation to get back into the deck — a font face that may not exist on
 * the target machine, a highlight, a baseline shift — is out of surface on
 * purpose: the return path must never have to guess.
 */
export const EDITABLE_RUN_PROPS = ['bold', 'italic', 'sizePt', 'color'] as const

export type EditableRunProp = (typeof EDITABLE_RUN_PROPS)[number]

/** One rule of the surface. Paths are dotted, relative to a `RenderSlide`. */
export interface SurfaceRule {
	kind: 'text' | 'runProp' | 'delete'
	path: string
	why: string
}

const RUNS = 'nodes[].text.paragraphs[].runs[]'

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

export interface ProjectedNode {
	id: NodeId
	kind: NodeKind
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
 * This is what part 05 renders as editable regions and what part 06 reads back.
 * A difference between the projection that went out and the one that came back
 * is an *edit*; a difference anywhere else is *drift* — see {@link freeze}.
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

	const runs: ProjectedRun[] = []
	if (node.kind === 'shape') collectRuns(node.id, node.text, runs)
	if (node.kind === 'table') {
		for (const row of node.rows) {
			for (const cell of row.cells) collectRuns(cell.id, cell.text, runs)
		}
	}
	return [{ id: node.id, kind: node.kind, runs }]
}

/** The shape {@link collectRuns} needs; both `ShapeNode.text` and `TableCell.text` satisfy it. */
type TextBodyLike = { paragraphs: { runs: { text: string; props: RunProperties }[] }[] } | null

function collectRuns(owner: NodeId, text: TextBodyLike, into: ProjectedRun[]): void {
	if (!text) return
	text.paragraphs.forEach((paragraph, paragraphIndex) => {
		paragraph.runs.forEach((run, runIndex) => {
			into.push({
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
 * (`data-d2p-props`) and the return path reads exactly this back. That attribute
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

// ---------------------------------------------------------------------------
// The complement
// ---------------------------------------------------------------------------

/**
 * The IR with every editable value stripped — the half that is *not* allowed to
 * change.
 *
 * Part 06 hashes this. Hashing the whole IR would flag a legitimate text edit as
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
	// The IR is JSON-safe by contract (see `./render`), so this is a deep clone.
	const clone = JSON.parse(JSON.stringify(ir)) as RenderIr
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
		for (const run of paragraph.runs) {
			run.text = ''
			for (const key of EDITABLE_RUN_PROPS) delete run.props[key]
		}
	}
}
