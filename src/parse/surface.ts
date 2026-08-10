/**
 * Reading the editable surface back out of a document.
 *
 * This is the only place in the return path that looks at the visual DOM, and it
 * looks at exactly three things: the text inside each `[data-pxh-run]` span, the
 * `data-pxh-props` attribute beside it, and the `data-pxh-paraprops` on the
 * `[data-pxh-para]` that encloses them. Nothing else. It does not read a computed
 * style, a transform, a path or a fill, because the island already states all of
 * those *better* than the DOM can — and re-deriving them is the inference this
 * architecture exists to avoid.
 *
 * ## The island is the schema
 *
 * The reading is driven by `project(ir)` rather than by a DOM query, because a
 * projection built from the DOM alone would be missing things the DOM never
 * carried: a node's `kind`, and the distinction between a group (which
 * contributes no node of its own) and a table cell (whose runs belong to the
 * table). The island says what *should* be there; the DOM says what it now says.
 * That asymmetry is also what makes an absent element mean something specific
 * rather than being invisible.
 *
 * ## What drift this can and cannot see
 *
 * The plan's drifted lane is "the DOM differs outside the surface". Detected in
 * full: a run or paragraph element removed, an address that is not in the model, a
 * duplicate address, and a `data-pxh-props` / `data-pxh-paraprops` that is not a
 * well-formed set of the editable properties. **Not** detected: a moved box, a
 * recoloured path, a rewritten transform.
 *
 * That gap is deliberate and it is safe, for a reason worth stating rather than
 * apologising for: **an out-of-surface DOM edit is inert, not dangerous.** The
 * emit path reads the island and this reading, and nothing else — so a shape a
 * user dragged in dev tools is not misapplied, it is simply not applied. The loss
 * is the user's edit, never the deck's fidelity. Detecting it would mean
 * re-deriving the model from the DOM to compare, which is the one thing that
 * would put inference back into the trusted path.
 */

import type { NodeId, ParagraphProperties, RenderIr, RunProperties } from '../ir/render'
import {
	EDITABLE_PARA_PROPS,
	EDITABLE_RUN_PROPS,
	type EditableParaProp,
	type EditableRunProp,
	editableParaProps,
	editableRunProps,
	type ProjectedParagraph,
	type ProjectedRun,
	project,
	type SanctionedProjection,
} from '../ir/surface'

/** What the DOM said, in the same shape the island states it. */
export interface SurfaceReading {
	projection: SanctionedProjection
	/** Nodes the renderer drew and the document no longer has — a sanctioned deletion. */
	deleted: NodeId[]
	/**
	 * Edits the surface does not define. Each one is kept as *the island's* value
	 * and reported; none of them is ever interpreted.
	 */
	anomalies: string[]
}

/** `data-pxh-props`, validated. Anything unexpected is refused rather than coerced. */
function readProps(raw: string | null, address: string, anomalies: string[]): Pick<RunProperties, EditableRunProp> {
	if (raw === null) return {}
	let parsed: unknown
	try {
		parsed = JSON.parse(raw)
	} catch {
		anomalies.push(`run ${address}: data-pxh-props is not JSON; the run's stated formatting was kept`)
		return {}
	}
	if (parsed === null || typeof parsed !== 'object' || Array.isArray(parsed)) {
		anomalies.push(`run ${address}: data-pxh-props is not an object; the run's stated formatting was kept`)
		return {}
	}

	const entries = parsed as Record<string, unknown>
	for (const key of Object.keys(entries)) {
		if (!(EDITABLE_RUN_PROPS as readonly string[]).includes(key)) {
			anomalies.push(`run ${address}: data-pxh-props carries ${JSON.stringify(key)}, which is not in the editable surface`)
		}
	}

	const picked: Record<string, unknown> = {}
	for (const key of EDITABLE_RUN_PROPS) {
		const value = entries[key]
		if (value === undefined) continue
		const complaint = valueComplaint(key, value)
		if (complaint !== null) {
			anomalies.push(`run ${address}: data-pxh-props.${key} ${complaint}`)
			continue
		}
		picked[key] = value
	}
	// Back through the same narrowing the renderer wrote it with, so an unedited
	// document's projection is not merely equal but *identically serialized* — the
	// two hashes are compared as strings and key order is part of a JSON string.
	return editableRunProps(picked as RunProperties)
}

/**
 * `data-pxh-paraprops`, validated the same way and by the same rules.
 *
 * Kept as its own function rather than folded into {@link readProps} with a
 * parameterised key list: the two attributes carry different property sets and
 * different value shapes, and a shared reader would have to be keyed off a union,
 * which is precisely the sort of "one of the two" indirection that lets a value
 * from one tier be accepted on the other.
 */
function readParaProps(
	raw: string | null,
	address: string,
	anomalies: string[]
): Pick<ParagraphProperties, EditableParaProp> {
	if (raw === null) return {}
	let parsed: unknown
	try {
		parsed = JSON.parse(raw)
	} catch {
		anomalies.push(`paragraph ${address}: data-pxh-paraprops is not JSON; the paragraph's stated formatting was kept`)
		return {}
	}
	if (parsed === null || typeof parsed !== 'object' || Array.isArray(parsed)) {
		anomalies.push(`paragraph ${address}: data-pxh-paraprops is not an object; the paragraph's stated formatting was kept`)
		return {}
	}

	const entries = parsed as Record<string, unknown>
	for (const key of Object.keys(entries)) {
		if (!(EDITABLE_PARA_PROPS as readonly string[]).includes(key)) {
			anomalies.push(
				`paragraph ${address}: data-pxh-paraprops carries ${JSON.stringify(key)}, which is not in the editable surface`
			)
		}
	}

	const picked: Record<string, unknown> = {}
	for (const key of EDITABLE_PARA_PROPS) {
		const value = entries[key]
		if (value === undefined) continue
		const complaint = paraValueComplaint(key, value)
		if (complaint !== null) {
			anomalies.push(`paragraph ${address}: data-pxh-paraprops.${key} ${complaint}`)
			continue
		}
		picked[key] = value
	}
	return editableParaProps(picked as Pick<ParagraphProperties, EditableParaProp>)
}

/** The four `a:pPr/@algn` values the write API expresses, and the only four `align` may be. */
const ALIGN_VALUES = new Set(['left', 'center', 'right', 'justify'])

/**
 * What is wrong with one stated paragraph property, or `null` if nothing is.
 *
 * Keyed off {@link EDITABLE_PARA_PROPS} for the same reason its run-level
 * counterpart is: a property added there without a case here must not compile.
 */
function paraValueComplaint(key: EditableParaProp, value: unknown): string | null {
	switch (key) {
		case 'align':
			// `dist` and `thaiDist` are real `ST_TextAlignType` members and are
			// deliberately not among them — import files a `text.align` note for those
			// rather than rounding one into a neighbour, so accepting one back here
			// would put a value into the model that emit has nowhere to send.
			return typeof value === 'string' && ALIGN_VALUES.has(value)
				? null
				: `is ${JSON.stringify(value)} and must be one of left, center, right, justify`
		case 'bullet':
			return bulletComplaint(value)
	}
}

/**
 * A bullet is a well-formed {@link Bullet}, and that is the whole of the test.
 *
 * All four kinds are accepted, which is a wider gate than `align`'s and is the
 * right one for a different reason. `align` refuses `dist` because import never
 * puts `dist` in the model, so a document stating one has been tampered with. A
 * numbered or picture bullet is the opposite: it is exactly what a real deck holds
 * and what the renderer just wrote into this attribute, so refusing it here would
 * report an *unedited* slide as drifted.
 *
 * Only three of those states can be authored *back* — see `PARA_INHERITED_OF` and
 * `paraOptionValue` in `parse/edits.ts` — and the gap is safe because nothing is
 * written unless it changed. This function's job is to keep a malformed value out
 * of the model, not to decide what may be set.
 */
function bulletComplaint(value: unknown): string | null {
	if (value === null || typeof value !== 'object' || Array.isArray(value)) {
		return `is ${typeof value} and must be a Bullet object`
	}
	const kind = (value as { kind?: unknown }).kind
	if (kind === 'none') return null
	if (kind === 'character') {
		return typeof (value as { char?: unknown }).char === 'string' ? null : 'is a character bullet stating no char'
	}
	if (kind === 'number') {
		return typeof (value as { scheme?: unknown }).scheme === 'string' ? null : 'is a numbered bullet stating no scheme'
	}
	if (kind === 'picture') {
		const asset = (value as { asset?: { $asset?: unknown } }).asset
		return typeof asset?.$asset === 'string' ? null : 'is a picture bullet stating no asset'
	}
	return `has kind ${JSON.stringify(kind)}, which is not none, character, number or picture`
}

/** The three values `underline` and `strike` each model, and the only three either may be. */
const DECORATION_VALUES = new Set(['none', 'single', 'double'])

/**
 * What is wrong with one stated property, or `null` if nothing is.
 *
 * Keyed off {@link EDITABLE_RUN_PROPS} so a property added there without a case
 * here is a type error rather than a value waved through unchecked — the whole
 * point of this function is that no surface value reaches the model unexamined.
 */
function valueComplaint(key: EditableRunProp, value: unknown): string | null {
	switch (key) {
		case 'color':
			return colorComplaint(value)
		case 'sizePt':
			return typeof value === 'number' ? null : `is ${typeof value} and must be number`
		case 'bold':
		case 'italic':
			return typeof value === 'boolean' ? null : `is ${typeof value} and must be boolean`
		case 'underline':
		case 'strike':
			// The enum is narrow on purpose: `RunProperties` models three of
			// `ST_TextUnderlineType`'s eighteen members because those are the three the
			// write API expresses, so accepting a fourth here would put a value into the
			// model that emit has nowhere to send.
			return typeof value === 'string' && DECORATION_VALUES.has(value)
				? null
				: `is ${JSON.stringify(value)} and must be one of none, single, double`
	}
}

/**
 * A colour is a {@link Color}, not a hex string.
 *
 * The write API takes `"250F6B"` or `"accent1"`, and it would be easy to let that
 * spelling leak up here — but the two arms mean different things on the way back:
 * a scheme colour keeps tracking the theme and an sRGB one does not. Flattening a
 * `scheme` into its resolved hex on the return path would bake the source theme
 * into every edited run, so this refuses the string form rather than accepting it
 * and guessing which arm it meant.
 */
function colorComplaint(value: unknown): string | null {
	if (value === null || typeof value !== 'object') return `is ${typeof value} and must be a Color object`
	const kind = (value as { kind?: unknown }).kind
	if (kind === 'srgb') {
		return typeof (value as { hex?: unknown }).hex === 'string' ? null : 'is an srgb Color with no hex'
	}
	if (kind === 'scheme') {
		return typeof (value as { slot?: unknown }).slot === 'string' ? null : 'is a scheme Color with no slot'
	}
	return `has kind ${JSON.stringify(kind)}, which is not srgb or scheme`
}

/**
 * Read the surface from a rendered document.
 *
 * `root` is a live `Document` in the browser, or anything `querySelector` reaches
 * — a `DOMParser` result, a detached subtree. There is no string form: the edits
 * this reads live in a DOM, and a string that was never one cannot have any.
 */
export function readSurface(root: ParentNode, ir: RenderIr): SurfaceReading {
	const island = project(ir)
	const anomalies: string[] = []
	const deleted: NodeId[] = []
	const claimed = new Set<string>()
	const claimedParagraphs = new Set<string>()

	// Which nodes were actually drawn. A node with no resolvable placement is not
	// in the document at all (see `render/node.ts`), so its absence is the
	// renderer's doing and not a deletion — reading it as one would silently drop a
	// shape from the deck on every round trip.
	const drawn = new Set<NodeId>()
	for (const slide of ir.slides) collectDrawn(slide.nodes, drawn)

	const slides = island.slides.map((slide) => ({
		number: slide.number,
		nodes: slide.nodes.flatMap((node) => {
			const element = root.querySelector(`[data-pxh-node="${cssEscape(node.id)}"]`)
			if (element === null) {
				if (drawn.has(node.id)) {
					deleted.push(node.id)
					return []
				}
				// Never drawn, so never deletable: keep what the island says.
				return [node]
			}
			return [
				{
					...node,
					paragraphs: node.paragraphs.map((paragraph) =>
						readParagraph(root, paragraph, claimedParagraphs, anomalies)
					),
					runs: node.runs.map((run) => readRun(root, run, claimed, anomalies)),
				},
			]
		}),
	}))

	for (const element of root.querySelectorAll('[data-pxh-run]')) {
		const address = element.getAttribute('data-pxh-run') ?? ''
		if (!claimed.has(address)) {
			anomalies.push(`run ${address} is in the document but not in the model; it was ignored`)
		}
	}
	for (const element of root.querySelectorAll('[data-pxh-para]')) {
		const address = element.getAttribute('data-pxh-para') ?? ''
		if (!claimedParagraphs.has(address)) {
			anomalies.push(`paragraph ${address} is in the document but not in the model; it was ignored`)
		}
	}

	return { projection: { irVersion: island.irVersion, slides }, deleted, anomalies }
}

function collectDrawn(nodes: RenderIr['slides'][number]['nodes'], into: Set<NodeId>): void {
	for (const node of nodes) {
		if (node.placement === null) continue
		into.add(node.id)
		if (node.kind === 'group') collectDrawn(node.children, into)
		if (node.kind === 'table') {
			for (const row of node.rows) for (const cell of row.cells) into.add(cell.id)
		}
	}
}

/**
 * One paragraph, read back.
 *
 * The mirror of {@link readRun} minus the text: a paragraph holds no string of its
 * own, so the only thing to read is the attribute — and a paragraph whose element
 * is gone keeps what the island said, because removing a `<p>` is a structural
 * change the surface has no rule for.
 */
function readParagraph(
	root: ParentNode,
	paragraph: ProjectedParagraph,
	claimed: Set<string>,
	anomalies: string[]
): ProjectedParagraph {
	const address = `${paragraph.node}/${paragraph.paragraph}`
	const matches = root.querySelectorAll(`[data-pxh-para="${cssEscape(address)}"]`)
	if (matches.length === 0) {
		anomalies.push(`paragraph ${address} was removed from the document; the model's formatting was kept`)
		return paragraph
	}
	claimed.add(address)
	if (matches.length > 1) {
		anomalies.push(`paragraph ${address} appears ${matches.length} times in the document; the first was read`)
	}

	const element = matches[0] as Element
	return {
		node: paragraph.node,
		paragraph: paragraph.paragraph,
		props: readParaProps(element.getAttribute('data-pxh-paraprops'), address, anomalies),
	}
}

function readRun(root: ParentNode, run: ProjectedRun, claimed: Set<string>, anomalies: string[]): ProjectedRun {
	const address = `${run.node}/${run.paragraph}/${run.run}`
	const matches = root.querySelectorAll(`[data-pxh-run="${cssEscape(address)}"]`)
	if (matches.length === 0) {
		// The run's node survived but the run itself did not. Emptying a run is an
		// edit; removing its element is a structural change the surface has no rule
		// for, so the island's text stands.
		anomalies.push(`run ${address} was removed from the document; the model's text was kept`)
		return run
	}
	claimed.add(address)
	if (matches.length > 1) {
		anomalies.push(`run ${address} appears ${matches.length} times in the document; the first was read`)
	}

	const element = matches[0] as Element
	return {
		node: run.node,
		paragraph: run.paragraph,
		run: run.run,
		text: element.textContent ?? '',
		props: readProps(element.getAttribute('data-pxh-props'), address, anomalies),
	}
}

/**
 * Node ids are structural (`s1.sp7`, `s1.sp7.r0c1`) and cannot contain either of
 * these characters today. This is here so that stops being load-bearing: the day
 * an id derivation changes, the failure should be a selector that finds nothing,
 * not one that matches the wrong element. `CSS.escape` is browser-only and this
 * runs in Node too.
 */
function cssEscape(value: string): string {
	return value.replaceAll('\\', '\\\\').replaceAll('"', '\\"')
}
