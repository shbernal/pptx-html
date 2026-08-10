/**
 * Reading the editable surface out of a real document, in a real browser.
 *
 * The rest of the return path is proved in Node — `reconcile` is pure, the island
 * is inert text — and this file covers the one step that cannot be: whether the
 * DOM the renderer produced says the same thing the island does, once a browser
 * has parsed it, normalized its attributes and applied its own idea of
 * whitespace. jsdom would answer a slightly different question; the render target
 * is a browser, so the test is in one.
 *
 * The load-bearing claim is the first one. If an *untouched* document does not
 * read back as the projection that produced it, every document is reported as
 * edited, the exact lane never fires, and the whole four-lane scheme collapses
 * into "everything is reconciled" without a single test going red anywhere else.
 */

import { describe, expect, it } from 'vitest'
import type { RenderIr } from '../../src/ir/render'
import { project } from '../../src/ir/surface'
import { blocksOf } from '../../src/parse/island'
import { reconcile } from '../../src/parse/reconcile'
import { readSurface } from '../../src/parse/surface'
import { renderDeck } from '../../src/render/document'
import { SAMPLE_IR } from '../fixtures/render-ir'

const NO_ASSETS = { assets: 'ref' } as const

async function documentOf(ir: RenderIr): Promise<Document> {
	const { html } = await renderDeck(ir, NO_ASSETS)
	return new DOMParser().parseFromString(html, 'text/html')
}

describe('an untouched document', () => {
	it('reads back as exactly the projection it was rendered from', async () => {
		const dom = await documentOf(SAMPLE_IR)
		const reading = readSurface(dom, SAMPLE_IR)
		expect(reading.anomalies).toEqual([])
		expect(reading.deleted).toEqual([])
		// Serialized, not deep-equal: `surfaceHash` is taken over the JSON string, so
		// key order is part of the comparison the lane decision actually makes.
		expect(JSON.stringify(reading.projection)).toBe(JSON.stringify(project(SAMPLE_IR)))
	})

	it('takes the exact lane on every slide', async () => {
		const dom = await documentOf(SAMPLE_IR)
		const result = reconcile(SAMPLE_IR, readSurface(dom, SAMPLE_IR))
		expect(result.slides.every((slide) => slide.lane === 'exact')).toBe(true)
	})

	it('reads the same island blocks as the string scanner', async () => {
		// The two readers exist because the island has to be legible in Node and in
		// the browser; if they ever disagree, one environment silently loses a deck.
		const { html } = await renderDeck(SAMPLE_IR, NO_ASSETS)
		expect(blocksOf(new DOMParser().parseFromString(html, 'text/html'))).toStrictEqual(blocksOf(html))
	})
})

describe('an edited document', () => {
	it('carries retyped text back and marks only that slide reconciled', async () => {
		const dom = await documentOf(SAMPLE_IR)
		const span = dom.querySelector('[data-pxh-run]')
		if (span === null) throw new Error('the rendered document has no editable runs')
		span.textContent = 'typed by a human'

		const result = reconcile(SAMPLE_IR, readSurface(dom, SAMPLE_IR))
		expect(result.slides[0]?.lane).toBe('reconciled')
		expect(JSON.stringify(result.ir)).toContain('typed by a human')
	})

	it('reads a changed character property from the attribute, not from the painted style', async () => {
		// The distinction that keeps a placeholder from being flattened: the span's
		// `style` is `props ?? resolved`, so reading a size off it would write the
		// layout's value into the slide. Only `data-pxh-props` is in surface.
		const dom = await documentOf(SAMPLE_IR)
		const span = dom.querySelector('[data-pxh-run]')
		if (span === null) throw new Error('the rendered document has no editable runs')
		span.setAttribute('data-pxh-props', JSON.stringify({ bold: true }))

		const address = span.getAttribute('data-pxh-run')
		const read = readSurface(dom, SAMPLE_IR)
			.projection.slides.flatMap((slide) => slide.nodes)
			.flatMap((node) => node.runs)
			.find((run) => `${run.node}/${run.paragraph}/${run.run}` === address)
		expect(read?.props).toStrictEqual({ bold: true })
	})

	it('reads underline and strike back and reconciles the slide rather than drifting it', async () => {
		// The lane is the assertion, not the value: these two arrived in the surface
		// after `bold` and `sizePt`, and a reader that did not know them would file
		// each as "carries a key that is not in the editable surface" — reporting a
		// sanctioned edit as drift, which is the failure mode that makes the signal
		// worthless by crying wolf.
		const dom = await documentOf(SAMPLE_IR)
		const span = dom.querySelector('[data-pxh-run]')
		if (span === null) throw new Error('the rendered document has no editable runs')
		span.setAttribute('data-pxh-props', JSON.stringify({ underline: 'double', strike: 'none' }))

		const reading = readSurface(dom, SAMPLE_IR)
		expect(reading.anomalies).toEqual([])

		const address = span.getAttribute('data-pxh-run')
		const read = reading.projection.slides
			.flatMap((slide) => slide.nodes)
			.flatMap((node) => node.runs)
			.find((run) => `${run.node}/${run.paragraph}/${run.run}` === address)
		expect(read?.props).toStrictEqual({ underline: 'double', strike: 'none' })

		const result = reconcile(SAMPLE_IR, reading)
		expect(result.slides[0]?.lane).toBe('reconciled')
	})

	it('refuses a decoration value outside the three the write API expresses', async () => {
		// `ST_TextUnderlineType` has eighteen members and `RunProperties` models
		// three, because three is what comes back. Accepting `wavyHeavy` here would
		// put a value into the model that emit has nowhere to send — lost at the far
		// end instead of refused at this one.
		const dom = await documentOf(SAMPLE_IR)
		const span = dom.querySelector('[data-pxh-run]')
		if (span === null) throw new Error('the rendered document has no editable runs')
		span.setAttribute('data-pxh-props', JSON.stringify({ underline: 'wavyHeavy' }))

		const reading = readSurface(dom, SAMPLE_IR)
		expect(reading.anomalies.join('\n')).toContain('must be one of none, single, double')
		expect(JSON.stringify(reconcile(SAMPLE_IR, reading).ir)).not.toContain('wavyHeavy')
	})

	it('reads a paragraph’s alignment off the `<p>`, not off the runs inside it', async () => {
		// The paragraph tier's version of the attribute test above, and the reason it
		// has to be here rather than in Node: a `<p>` is the one element in this
		// document a browser may reparent or normalize, and the address has to survive
		// that. `text-align` is also on the same element's `style`, so this is again
		// the distinction between what the deck *states* and what the browser paints.
		const dom = await documentOf(SAMPLE_IR)
		const para = dom.querySelector('[data-pxh-para]')
		if (para === null) throw new Error('the rendered document has no addressable paragraphs')
		para.setAttribute('data-pxh-paraprops', JSON.stringify({ align: 'justify' }))

		const reading = readSurface(dom, SAMPLE_IR)
		expect(reading.anomalies).toEqual([])

		const address = para.getAttribute('data-pxh-para')
		const read = reading.projection.slides
			.flatMap((slide) => slide.nodes)
			.flatMap((node) => node.paragraphs)
			.find((paragraph) => `${paragraph.node}/${paragraph.paragraph}` === address)
		expect(read?.props).toStrictEqual({ align: 'justify' })
		expect(reconcile(SAMPLE_IR, reading).slides[0]?.lane).toBe('reconciled')
	})

	it('refuses an alignment outside the four the write API expresses', async () => {
		// `dist` is a real `ST_TextAlignType` member that import files a note for
		// rather than rounding into a neighbour, so accepting it back here would put a
		// value into the model with nowhere to send it.
		const dom = await documentOf(SAMPLE_IR)
		const para = dom.querySelector('[data-pxh-para]')
		if (para === null) throw new Error('the rendered document has no addressable paragraphs')
		para.setAttribute('data-pxh-paraprops', JSON.stringify({ align: 'dist' }))

		const reading = readSurface(dom, SAMPLE_IR)
		expect(reading.anomalies.join('\n')).toContain('must be one of left, center, right, justify')
		expect(JSON.stringify(reconcile(SAMPLE_IR, reading).ir)).not.toContain('"dist"')
	})

	it('reads a bullet back off the `<p>`, glyph and all', async () => {
		// The second paragraph property, and the one whose value is an object rather
		// than a token — so this is also the test that the attribute survives a
		// round trip through `JSON.stringify` and the browser's own attribute handling
		// with its structure intact rather than flattened to a string.
		const dom = await documentOf(SAMPLE_IR)
		const para = dom.querySelector('[data-pxh-para]')
		if (para === null) throw new Error('the rendered document has no addressable paragraphs')
		para.setAttribute('data-pxh-paraprops', JSON.stringify({ bullet: { kind: 'character', char: '▸' } }))

		const reading = readSurface(dom, SAMPLE_IR)
		expect(reading.anomalies).toEqual([])

		const address = para.getAttribute('data-pxh-para')
		const read = reading.projection.slides
			.flatMap((slide) => slide.nodes)
			.flatMap((node) => node.paragraphs)
			.find((paragraph) => `${paragraph.node}/${paragraph.paragraph}` === address)
		expect(read?.props).toStrictEqual({ bullet: { kind: 'character', char: '▸' } })
		expect(reconcile(SAMPLE_IR, reading).slides[0]?.lane).toBe('reconciled')
	})

	it('refuses a bullet that is not one of the four kinds, and keeps the model’s', async () => {
		// The gate here is wider than `align`'s — all four kinds are accepted, because
		// a numbered or picture bullet is exactly what a real deck holds and refusing
		// one would report an unedited slide as drifted. What it still refuses is a
		// value that is not a `Bullet` at all, which no renderer of this package wrote.
		const dom = await documentOf(SAMPLE_IR)
		const para = dom.querySelector('[data-pxh-para]')
		if (para === null) throw new Error('the rendered document has no addressable paragraphs')
		para.setAttribute('data-pxh-paraprops', JSON.stringify({ bullet: { kind: 'emoji', char: '🔥' } }))

		const reading = readSurface(dom, SAMPLE_IR)
		expect(reading.anomalies.join('\n')).toContain('which is not none, character, number or picture')
		expect(JSON.stringify(reconcile(SAMPLE_IR, reading).ir)).not.toContain('"emoji"')
	})

	it('reads a paragraph’s margins back as numbers, and clearing one as absence', async () => {
		// The margins are the surface's only measurement on the paragraph tier, and the
		// attribute is text — so this is the test that a point value survives the
		// document as a number rather than as `"18"`, which `reconcile` would read as a
		// change on every pass and emit would hand the writer as a string.
		const dom = await documentOf(SAMPLE_IR)
		const para = dom.querySelector('[data-pxh-para]')
		if (para === null) throw new Error('the rendered document has no addressable paragraphs')
		para.setAttribute('data-pxh-paraprops', JSON.stringify({ marginLeftPt: 36 }))

		const reading = readSurface(dom, SAMPLE_IR)
		expect(reading.anomalies).toEqual([])

		const address = para.getAttribute('data-pxh-para')
		const read = reading.projection.slides
			.flatMap((slide) => slide.nodes)
			.flatMap((node) => node.paragraphs)
			.find((paragraph) => `${paragraph.node}/${paragraph.paragraph}` === address)
		// The alignment this paragraph stated is gone from the attribute, which is how a
		// document spells *back to inherited* — so the reading has to be the margin
		// alone rather than the margin folded into what the island still says.
		expect(read?.props).toStrictEqual({ marginLeftPt: 36 })
		expect(reconcile(SAMPLE_IR, reading).slides[0]?.lane).toBe('reconciled')
	})

	it('refuses a margin that is not a number, and keeps the model’s', async () => {
		// The shape gate, not the range one: `parse/edits.ts` decides what the attribute
		// can hold, because a value out of range is still a margin the file may state.
		// A string is not, and accepting one would hand the writer `"36pt"` to drop.
		const dom = await documentOf(SAMPLE_IR)
		const para = dom.querySelector('[data-pxh-para]')
		if (para === null) throw new Error('the rendered document has no addressable paragraphs')
		para.setAttribute('data-pxh-paraprops', JSON.stringify({ marginLeftPt: '36pt' }))

		const reading = readSurface(dom, SAMPLE_IR)
		expect(reading.anomalies.join('\n')).toContain('must be a finite number of points')
		expect(JSON.stringify(reconcile(SAMPLE_IR, reading).ir)).not.toContain('36pt')
	})

	it('does not address a chrome paragraph, so the template’s own text is not read as drift', async () => {
		// Chrome is drawn and nothing more. Its `<p>` carries no `data-pxh-para`, which
		// is what keeps the sweep for unmodeled paragraphs quiet — without it every
		// slide would report the layout's every paragraph as an anomaly, on every pass.
		const dom = await documentOf(SAMPLE_IR)
		const chrome = dom.querySelector('[data-pxh-chrome]')
		expect(chrome).not.toBeNull()
		expect(chrome?.querySelector('[data-pxh-para]')).toBeNull()
		expect(readSurface(dom, SAMPLE_IR).anomalies).toEqual([])
	})

	it('reports a removed node as a deletion and a removed run as drift', async () => {
		// Two structurally similar edits with opposite meanings. Deleting a node is in
		// surface and is honoured; deleting a *run* is not, so the model's text stands
		// and the slide is marked. Reading either as the other loses content silently.
		const dom = await documentOf(SAMPLE_IR)
		const doomedNode = SAMPLE_IR.slides[0]?.nodes[1]
		if (doomedNode === undefined) throw new Error('the fixture lost its second node')
		dom.querySelector(`[data-pxh-node="${doomedNode.id}"]`)?.remove()

		const first = readSurface(dom, SAMPLE_IR)
		expect(first.deleted).toContain(doomedNode.id)
		expect(first.anomalies).toEqual([])

		const survivor = dom.querySelector('[data-pxh-run]')
		if (survivor === null) throw new Error('the document has no runs left')
		const address = survivor.getAttribute('data-pxh-run')
		survivor.remove()

		const second = readSurface(dom, SAMPLE_IR)
		expect(second.anomalies.join('\n')).toContain(`run ${address} was removed`)
		expect(reconcile(SAMPLE_IR, second).slides[0]?.lane).toBe('drifted')
	})

	it('refuses a property the surface does not define instead of writing it', async () => {
		const dom = await documentOf(SAMPLE_IR)
		const span = dom.querySelector('[data-pxh-run]')
		if (span === null) throw new Error('the rendered document has no editable runs')
		span.setAttribute('data-pxh-props', JSON.stringify({ bold: true, fontFace: 'Comic Sans MS' }))

		const reading = readSurface(dom, SAMPLE_IR)
		expect(reading.anomalies.join('\n')).toContain('"fontFace"')
		expect(JSON.stringify(reconcile(SAMPLE_IR, reading).ir)).not.toContain('Comic Sans MS')
	})

	it('ignores a run address the model does not have', async () => {
		const dom = await documentOf(SAMPLE_IR)
		const span = dom.querySelector('[data-pxh-run]')
		if (span === null) throw new Error('the rendered document has no editable runs')
		const invented = span.cloneNode(true) as Element
		invented.setAttribute('data-pxh-run', 's1.sp999/0/0')
		span.parentElement?.append(invented)

		const reading = readSurface(dom, SAMPLE_IR)
		expect(reading.anomalies.join('\n')).toContain('s1.sp999/0/0 is in the document but not in the model')
	})
})
