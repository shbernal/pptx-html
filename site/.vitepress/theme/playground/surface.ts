/**
 * The editable surface, made operable.
 *
 * The panel this drives is `EDITABLE_SURFACE` with controls attached: run text,
 * `bold` / `italic` / `underline` / `strike` / `sizePt` / `color`, the paragraph's
 * `align`, `bullet` and two margins, and node deletion. Nothing else gets a
 * control, because an input whose value is silently dropped at emit is exactly the
 * failure this project is arranged against, reproduced in its own demo.
 *
 * That rule is why the bullet control has a position it will not let you *choose*.
 * A paragraph can hold a glyph the write API cannot author back — a numbering
 * scheme, a picture, a glyph with its own colour — and the honest thing to show is
 * that the deck states one and this panel is not offering to replace it, rather
 * than a control that reads as "no bullet" or one that quietly rounds it to a dot.
 *
 * The pairing *was* hand-wired, and that was a divergence risk the surface file
 * exists to prevent: `EDITABLE_RUN_PROPS` said which properties are in surface,
 * nothing generated a control from it, and a property added there would have
 * gained no control — a quiet failure, a missing input rather than a broken
 * build. {@link CONTROL_OF} closes it. It is a `Record<EditableRunProp, …>`, so
 * a property added to the surface and not given a control kind here is a *type*
 * error, and the panel renders from its entries rather than from a second list.
 *
 * **Every edit is written into the live document, never into the model.** The
 * renderer already marks each run `contenteditable` and addresses it with
 * `data-pxh-run`, so the panel and a visitor typing into the preview are the same
 * edit arriving by two routes — and both are read back by the same `parseDeck`
 * call. Mutating the parsed IR instead would produce a demo that proves the
 * return path works by never using it.
 */

import type {
	Bullet,
	Color,
	EditableParaProp,
	EditableRunProp,
	ParagraphProperties,
	RenderIr,
	RunProperties,
} from 'pptx-html'
import { project } from 'pptx-html'

/**
 * How each surface property is operated. Five kinds, not seven: `underline` and
 * `strike` are the same control because they are the same shape — three named
 * values plus *inherited*, which is what an absent key means — and `align` is a
 * fourth of the same family, four named values plus the same absence.
 */
export type ControlKind = 'toggle' | 'decoration' | 'number' | 'color' | 'align' | 'bullet' | 'points'

/**
 * Every property in the surface, with the control that drives it.
 *
 * The `Record<EditableRunProp, …>` is the whole point: this file cannot compile
 * while the surface names a property it does not. The panel iterates these
 * entries, so the list is also the render order.
 */
export const CONTROL_OF: Record<EditableRunProp, ControlKind> = {
	bold: 'toggle',
	italic: 'toggle',
	underline: 'decoration',
	strike: 'decoration',
	sizePt: 'number',
	color: 'color',
}

/** The panel's control list, in surface order, with the key each one writes. */
export const CONTROLS: readonly { prop: EditableRunProp; kind: ControlKind }[] = (
	Object.keys(CONTROL_OF) as EditableRunProp[]
).map((prop) => ({ prop, kind: CONTROL_OF[prop] }))

/**
 * The paragraph tier, under the same rule and for the same reason: a
 * `Record<EditableParaProp, …>`, so a property added to the paragraph surface and
 * not given a control here is a `site:typecheck` error rather than a missing input
 * nobody notices.
 */
export const PARA_CONTROL_OF: Record<EditableParaProp, ControlKind> = {
	align: 'align',
	bullet: 'bullet',
	marginLeftPt: 'points',
	indentPt: 'points',
}

/**
 * The two margin inputs, with the range `a:pPr/@marL` and `@indent` each accept.
 *
 * `min` is not decoration: a negative `marL` is a value PowerPoint reports as
 * needing repair, and `parse/edits.ts` refuses one rather than letting the writer
 * clamp it. The input says so before the refusal has to.
 */
export const POINTS_BOUNDS: Record<'marginLeftPt' | 'indentPt', { min: number; max: number }> = {
	marginLeftPt: { min: 0, max: 4032 },
	indentPt: { min: -4032, max: 4032 },
}

export const PARA_CONTROLS: readonly { prop: EditableParaProp; kind: ControlKind }[] = (
	Object.keys(PARA_CONTROL_OF) as EditableParaProp[]
).map((prop) => ({ prop, kind: PARA_CONTROL_OF[prop] }))

/**
 * The four values `align` models, plus the absence that means *inherited* — the
 * same shape as {@link DECORATION_CHOICES}, and the same empty-string stand-in for
 * an absence a `<select>` cannot hold.
 */
export const ALIGN_CHOICES: readonly { value: string; label: string }[] = [
	{ value: '', label: 'inherited' },
	{ value: 'left', label: 'left' },
	{ value: 'center', label: 'center' },
	{ value: 'right', label: 'right' },
	{ value: 'justify', label: 'justify' },
]

/**
 * The bullet control's positions: the three states the write API can author, plus
 * one for a glyph it cannot.
 *
 * `STATED` is rendered `disabled` and is selected only when the paragraph already
 * holds such a bullet. It is the panel's way of saying *the deck states one and
 * this control will not touch it* — the alternative would be showing "none" for a
 * paragraph that visibly has a bullet, or offering a dot that would replace a
 * numbering scheme with a lie.
 */
export const BULLET_STATED = 'as-stated'

export const BULLET_CHOICES: readonly { value: string; label: string }[] = [
	{ value: '', label: 'inherited' },
	{ value: 'none', label: 'none' },
	{ value: 'bullet', label: '• bullet' },
	{ value: BULLET_STATED, label: 'as the deck states it' },
]

/** Which position a paragraph's bullet sits at. */
export function bulletChoiceOf(bullet: Bullet | undefined): string {
	if (bullet === undefined) return ''
	if (bullet.kind === 'none') return 'none'
	const plain =
		bullet.kind === 'character' &&
		bullet.char === '•' &&
		bullet.font === undefined &&
		bullet.color === undefined &&
		bullet.sizePct === undefined
	return plain ? 'bullet' : BULLET_STATED
}

/** And what that position means as a value. `undefined` clears the key, as everywhere else. */
export function bulletValueOf(choice: string): Bullet | undefined {
	if (choice === 'none') return { kind: 'none' }
	if (choice === 'bullet') return { kind: 'character', char: '•' }
	return undefined
}

/**
 * The three values `underline` and `strike` model, plus the absence that means
 * *inherited*. The empty string is the `<option>` value for that absence, since
 * a `<select>` has no way to hold `undefined`.
 */
export const DECORATION_CHOICES: readonly { value: string; label: string }[] = [
	{ value: '', label: 'inherited' },
	{ value: 'none', label: 'none' },
	{ value: 'single', label: 'single' },
	{ value: 'double', label: 'double' },
]

export interface RunRow {
	/** `node/paragraph/run` — the address the renderer wrote and the parser reads. */
	address: string
	/** Which paragraph holds it, so the panel can nest the two tiers without re-parsing the address. */
	paragraph: number
	text: string
	/** Exactly what the run *states*, which is what the projection carries. */
	props: Pick<RunProperties, EditableRunProp>
}

export interface ParagraphRow {
	/** `node/paragraph` — one address shorter than a run's, and the same idea. */
	address: string
	props: Pick<ParagraphProperties, EditableParaProp>
	/** The runs inside it, so the panel can show a paragraph as the box that holds them. */
	runs: RunRow[]
}

export interface NodeRow {
	id: string
	kind: string
	paragraphs: ParagraphRow[]
	runs: RunRow[]
}

export interface SlideRow {
	number: number
	nodes: NodeRow[]
}

/** The current state of the surface, read from the model the return path produced. */
export function surfaceOf(ir: RenderIr): SlideRow[] {
	return project(ir).slides.map((slide) => ({
		number: slide.number,
		nodes: slide.nodes.map((node) => {
			const runs = node.runs.map((run) => ({
				address: `${run.node}/${run.paragraph}/${run.run}`,
				text: run.text,
				props: run.props,
				paragraph: run.paragraph,
			}))
			return {
				id: node.id,
				kind: node.kind,
				paragraphs: node.paragraphs.map((paragraph) => ({
					address: `${paragraph.node}/${paragraph.paragraph}`,
					props: paragraph.props,
					runs: runs.filter((run) => run.paragraph === paragraph.paragraph),
				})),
				runs,
			}
		}),
	}))
}

export function setText(doc: Document, address: string, text: string): void {
	span(doc, address).textContent = text
}

/**
 * Set or clear one of the stated properties.
 *
 * `undefined` removes the key rather than writing `undefined`: absence means
 * *inherited*, and the two are different facts about the run. The reader refuses
 * anything that is not a well-formed member of the surface, so writing a
 * placeholder here would be reported as an anomaly rather than accepted.
 */
export function setProp(doc: Document, address: string, prop: EditableRunProp, value: unknown): void {
	const element = span(doc, address)
	const stated = statedProps(element, 'data-pxh-props')
	if (value === undefined) delete stated[prop]
	else stated[prop] = value

	if (Object.keys(stated).length === 0) element.removeAttribute('data-pxh-props')
	else element.setAttribute('data-pxh-props', JSON.stringify(stated))
}

/**
 * The same, one tier up: the `<p>` carries `data-pxh-paraprops` and nothing else
 * about it differs. Setting it on the paragraph rather than on each of its runs is
 * not a convenience — which runs of a paragraph a property has to land on to mean
 * what it says is a write-contract question, it differs per property, and it is
 * `parse/edits.ts`'s to answer. The document states the fact once.
 */
export function setParaProp(doc: Document, address: string, prop: EditableParaProp, value: unknown): void {
	const element = paragraph(doc, address)
	const stated = statedProps(element, 'data-pxh-paraprops')
	if (value === undefined) delete stated[prop]
	else stated[prop] = value

	if (Object.keys(stated).length === 0) element.removeAttribute('data-pxh-paraprops')
	else element.setAttribute('data-pxh-paraprops', JSON.stringify(stated))
}

/**
 * A colour is a `Color`, never a hex string.
 *
 * The write API accepts `"250F6B"`, and letting that spelling through here would
 * be read as a malformed property and refused — correctly, because the string
 * form cannot say whether it meant an sRGB value or a scheme reference.
 */
export function srgb(hex: string): Color {
	return { kind: 'srgb', hex: hex.replace(/^#/, '').toUpperCase() }
}

/** Delete a node. Its element leaving the document is how a deletion is spelled. */
export function deleteNode(doc: Document, id: string): void {
	const element = doc.querySelector(`[data-pxh-node="${cssEscape(id)}"]`)
	if (element === null) throw new Error(`node ${id} is not in the document`)
	element.remove()
}

/**
 * Typed as `Element`, and checked for `null` rather than with `instanceof`.
 *
 * The document lives in an iframe, so its elements come from that realm's
 * constructors — `element instanceof HTMLElement` is false against the parent
 * window's `HTMLElement` even when the element is exactly what it looks like.
 * Everything used here (`textContent`, the attribute methods) is on `Element`
 * anyway.
 */
function span(doc: Document, address: string): Element {
	const element = doc.querySelector(`[data-pxh-run="${cssEscape(address)}"]`)
	if (element === null) throw new Error(`run ${address} is not in the document`)
	return element
}

function paragraph(doc: Document, address: string): Element {
	const element = doc.querySelector(`[data-pxh-para="${cssEscape(address)}"]`)
	if (element === null) throw new Error(`paragraph ${address} is not in the document`)
	return element
}

function statedProps(element: Element, attribute: string): Record<string, unknown> {
	const raw = element.getAttribute(attribute)
	if (raw === null) return {}
	const parsed: unknown = JSON.parse(raw)
	if (parsed === null || typeof parsed !== 'object') throw new Error(`${attribute} is not an object`)
	return { ...(parsed as Record<string, unknown>) }
}

/**
 * The same escaping `src/parse/surface.ts` uses, for the same reason: these are
 * quoted attribute values, so a backslash and a double quote are the only two
 * characters that need it. `CSS.escape` is for identifiers, not strings.
 */
function cssEscape(value: string): string {
	return value.replaceAll('\\', '\\\\').replaceAll('"', '\\"')
}
