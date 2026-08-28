/**
 * Folding a document's edits back into the model.
 *
 * The plan's own warning about this step is worth repeating at the top of the
 * file it applies to: **this is where the guarantee usually dies.** Every
 * "helpful" interpretation of an edit is an inference, and inference is what the
 * whole architecture is arranged to keep out of the trusted path. So the rule
 * here has no exceptions — when the surface does not define what an edit means,
 * refuse it, keep the island's value, and say so.
 *
 * What that leaves is small and total: a run's text, the six character properties
 * `EDITABLE_RUN_PROPS` names, the four paragraph properties `EDITABLE_PARA_PROPS`
 * does, and whether a node still exists. Those are the only things this file
 * writes.
 *
 * ## The lanes
 *
 * A lane is decided per slide, and it is a fact about *that slide*, not about the
 * document — a deck where one slide was edited and eleven were not should not
 * describe all twelve the same way.
 *
 * | Lane | What happened |
 * |---|---|
 * | `exact` | the slide's surface came back identical; Invariant R holds outright |
 * | `reconciled` | a sanctioned edit; Invariant R holds for everything else |
 * | `drifted` | something outside the surface changed; the island's value stands |
 * | `heuristic` | there was no island at all — decided upstream of this file |
 *
 * `drifted` and `reconciled` are not exclusive: a slide can have a real text edit
 * *and* a run element someone deleted. It is reported as `drifted`, because that
 * is the outcome a caller has to act on.
 */

import { cloneIr, type NodeId, type ParagraphProperties, type RenderIr, type RenderNode, type RunProperties, type TextBody } from '../ir/render'
import {
	EDITABLE_PARA_PROPS,
	EDITABLE_RUN_PROPS,
	type ProjectedParagraph,
	type ProjectedRun,
	project,
	sameSurfaceValue,
} from '../ir/surface'
import type { SurfaceReading } from './surface'

/** How much of Invariant R a slide still carries. */
export type Lane = 'exact' | 'reconciled' | 'drifted' | 'heuristic'

export interface SlideOutcome {
	number: number
	lane: Lane
	/** Sanctioned edits applied to this slide: run text, run properties, node deletions. */
	edits: number
	/** Why the slide drifted, when it did. Empty otherwise. */
	notes: string[]
}

export interface Reconciliation {
	/** The island's model with the sanctioned edits folded in. */
	ir: RenderIr
	slides: SlideOutcome[]
	warnings: string[]
}

/**
 * Apply a surface reading to the model it was read against.
 *
 * Pure, and deliberately so: the DOM work is all in `surface.ts`, so the decision
 * logic — which is the part with a guarantee attached — is testable without a
 * browser, and the browser tests only have to prove that the reading is faithful.
 */
export function reconcile(ir: RenderIr, reading: SurfaceReading): Reconciliation {
	// The island's own model is never mutated: it is what `modelHash` covers, and a
	// caller comparing the two afterwards must be able to.
	const next = cloneIr(ir)
	const warnings: string[] = []

	const deleted = new Set(reading.deleted)
	const readRuns = new Map<string, ProjectedRun>()
	const readParagraphs = new Map<string, ProjectedParagraph>()
	for (const slide of reading.projection.slides) {
		for (const node of slide.nodes) {
			for (const paragraph of node.paragraphs) {
				readParagraphs.set(`${paragraph.node}/${paragraph.paragraph}`, paragraph)
			}
			for (const run of node.runs) readRuns.set(`${run.node}/${run.paragraph}/${run.run}`, run)
		}
	}

	// Anomalies are addressed by run or by node id, and both start with `s{n}.`, so
	// the slide they belong to is readable off the address. One that is not gets
	// attached to the deck rather than dropped.
	const anomaliesBySlide = new Map<number, string[]>()
	for (const anomaly of reading.anomalies) {
		const slide = slideNumberIn(anomaly)
		if (slide === null) warnings.push(anomaly)
		else anomaliesBySlide.set(slide, [...(anomaliesBySlide.get(slide) ?? []), anomaly])
	}

	const slides: SlideOutcome[] = next.slides.map((slide) => {
		let edits = 0
		const kept = slide.nodes.filter((node) => {
			const removed = pruneDeleted(node, deleted)
			edits += removed
			return !deleted.has(node.id)
		})
		edits += slide.nodes.length - kept.length
		slide.nodes = kept

		for (const node of slide.nodes) edits += applyRuns(node, readRuns, readParagraphs)

		const notes = anomaliesBySlide.get(slide.number) ?? []
		const lane: Lane = notes.length > 0 ? 'drifted' : edits > 0 ? 'reconciled' : 'exact'
		return { number: slide.number, lane, edits, notes }
	})

	return { ir: next, slides, warnings }
}

/**
 * Drop deleted descendants of a group, returning how many went.
 *
 * A group whose children were all deleted stays, and that is the conservative
 * reading rather than an oversight: the group is a node in its own right, an
 * empty one is what the source deck would hold if its shapes were removed one by
 * one, and inferring "you meant to delete the group too" is exactly the kind of
 * help this file refuses to give.
 */
function pruneDeleted(node: RenderNode, deleted: ReadonlySet<NodeId>): number {
	if (node.kind !== 'group') return 0
	let removed = 0
	for (const child of node.children) removed += pruneDeleted(child, deleted)
	const before = node.children.length
	node.children = node.children.filter((child) => !deleted.has(child.id))
	return removed + (before - node.children.length)
}

function applyRuns(
	node: RenderNode,
	read: ReadonlyMap<string, ProjectedRun>,
	readParagraphs: ReadonlyMap<string, ProjectedParagraph>
): number {
	if (node.kind === 'group') {
		return node.children.reduce((sum, child) => sum + applyRuns(child, read, readParagraphs), 0)
	}
	if (node.kind === 'shape') return applyText(node.id, node.text, read, readParagraphs)
	if (node.kind === 'table') {
		let edits = 0
		for (const row of node.rows) {
			for (const cell of row.cells) edits += applyText(cell.id, cell.text, read, readParagraphs)
		}
		return edits
	}
	return 0
}

function applyText(
	owner: NodeId,
	text: TextBody | null,
	read: ReadonlyMap<string, ProjectedRun>,
	readParagraphs: ReadonlyMap<string, ProjectedParagraph>
): number {
	if (text === null) return 0
	let edits = 0
	text.paragraphs.forEach((paragraph, paragraphIndex) => {
		const incomingParagraph = readParagraphs.get(`${owner}/${paragraphIndex}`)
		if (incomingParagraph !== undefined) {
			edits += applyProps(paragraph.props, incomingParagraph.props, EDITABLE_PARA_PROPS)
		}
		paragraph.runs.forEach((run, runIndex) => {
			const incoming = read.get(`${owner}/${paragraphIndex}/${runIndex}`)
			if (incoming === undefined) return
			if (incoming.text !== run.text) {
				run.text = incoming.text
				edits++
			}
			edits += applyProps(run.props, incoming.props, EDITABLE_RUN_PROPS)
		})
	})
	return edits
}

/**
 * Set the editable properties to exactly what came back.
 *
 * Absent has to mean *deleted*, not *unchanged*: absence is how this model spells
 * "inherited", so a run whose explicit bold was cleared in the document must lose
 * the key rather than keep it. Treating absence as "no news" would make clearing
 * a property the one edit the surface silently ignores.
 *
 * The key list is a parameter rather than a second copy of this function, and it
 * is the surface's own list either way — a run's properties and a paragraph's
 * differ in which keys they hold, never in what absence means.
 */
function applyProps<Target extends RunProperties | ParagraphProperties, Key extends keyof Target & string>(
	target: Target,
	incoming: Partial<Pick<Target, Key>>,
	keys: readonly Key[]
): number {
	let edits = 0
	for (const key of keys) {
		const value = incoming[key]
		if (value === undefined) {
			if (target[key] !== undefined) {
				delete target[key]
				edits++
			}
			continue
		}
		// Structurally, not by reference. `color` is an object, and the projection
		// arrived through a clone or a JSON parse, so `!==` is true for every run
		// that states one — which would report an untouched deck as fully edited.
		if (!sameSurfaceValue(target[key], value)) {
			Object.assign(target, { [key]: value })
			edits++
		}
	}
	return edits
}

/** `s4.sp7/0/1` → 4. `null` when the address is not one this model would produce. */
function slideNumberIn(text: string): number | null {
	const match = /\bs(\d+)\.(?:sp|dom)/.exec(text)
	const digits = match?.[1]
	return digits === undefined ? null : Number(digits)
}

/**
 * Whether a reading changed anything at all, without applying it.
 *
 * The same comparison the lane decision makes, over the serialized projections —
 * which is what {@link Integrity.surfaceHash} is a hash *of*, so this and the hash
 * can never disagree about whether a document was edited.
 */
export function surfaceChanged(ir: RenderIr, reading: SurfaceReading): boolean {
	return JSON.stringify(project(ir)) !== JSON.stringify(reading.projection)
}
