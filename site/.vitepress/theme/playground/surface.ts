/**
 * The editable surface, made operable.
 *
 * The panel this drives is `EDITABLE_SURFACE` with controls attached: run text,
 * `bold` / `italic` / `underline` / `strike` / `sizePt` / `color`, and node
 * deletion. Nothing else gets a control, because an input whose value is silently
 * dropped at emit is exactly the failure this project is arranged against,
 * reproduced in its own demo.
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

import type { Color, EditableRunProp, RenderIr, RunProperties } from 'pptx-html'
import { project } from 'pptx-html'

/**
 * How each surface property is operated. Four kinds, not six: `underline` and
 * `strike` are the same control because they are the same shape — three named
 * values plus *inherited*, which is what an absent key means.
 */
export type ControlKind = 'toggle' | 'decoration' | 'number' | 'color'

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
	text: string
	/** Exactly what the run *states*, which is what the projection carries. */
	props: Pick<RunProperties, EditableRunProp>
}

export interface NodeRow {
	id: string
	kind: string
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
		nodes: slide.nodes.map((node) => ({
			id: node.id,
			kind: node.kind,
			runs: node.runs.map((run) => ({
				address: `${run.node}/${run.paragraph}/${run.run}`,
				text: run.text,
				props: run.props,
			})),
		})),
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
	const stated = statedProps(element)
	if (value === undefined) delete stated[prop]
	else stated[prop] = value

	if (Object.keys(stated).length === 0) element.removeAttribute('data-pxh-props')
	else element.setAttribute('data-pxh-props', JSON.stringify(stated))
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

function statedProps(element: Element): Record<string, unknown> {
	const raw = element.getAttribute('data-pxh-props')
	if (raw === null) return {}
	const parsed: unknown = JSON.parse(raw)
	if (parsed === null || typeof parsed !== 'object') throw new Error('data-pxh-props is not an object')
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
