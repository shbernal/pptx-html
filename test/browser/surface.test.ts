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
		const span = dom.querySelector('[data-d2p-run]')
		if (span === null) throw new Error('the rendered document has no editable runs')
		span.textContent = 'typed by a human'

		const result = reconcile(SAMPLE_IR, readSurface(dom, SAMPLE_IR))
		expect(result.slides[0]?.lane).toBe('reconciled')
		expect(JSON.stringify(result.ir)).toContain('typed by a human')
	})

	it('reads a changed character property from the attribute, not from the painted style', async () => {
		// The distinction that keeps a placeholder from being flattened: the span's
		// `style` is `props ?? resolved`, so reading a size off it would write the
		// layout's value into the slide. Only `data-d2p-props` is in surface.
		const dom = await documentOf(SAMPLE_IR)
		const span = dom.querySelector('[data-d2p-run]')
		if (span === null) throw new Error('the rendered document has no editable runs')
		span.setAttribute('data-d2p-props', JSON.stringify({ bold: true }))

		const address = span.getAttribute('data-d2p-run')
		const read = readSurface(dom, SAMPLE_IR)
			.projection.slides.flatMap((slide) => slide.nodes)
			.flatMap((node) => node.runs)
			.find((run) => `${run.node}/${run.paragraph}/${run.run}` === address)
		expect(read?.props).toStrictEqual({ bold: true })
	})

	it('reports a removed node as a deletion and a removed run as drift', async () => {
		// Two structurally similar edits with opposite meanings. Deleting a node is in
		// surface and is honoured; deleting a *run* is not, so the model's text stands
		// and the slide is marked. Reading either as the other loses content silently.
		const dom = await documentOf(SAMPLE_IR)
		const doomedNode = SAMPLE_IR.slides[0]?.nodes[1]
		if (doomedNode === undefined) throw new Error('the fixture lost its second node')
		dom.querySelector(`[data-d2p-node="${doomedNode.id}"]`)?.remove()

		const first = readSurface(dom, SAMPLE_IR)
		expect(first.deleted).toContain(doomedNode.id)
		expect(first.anomalies).toEqual([])

		const survivor = dom.querySelector('[data-d2p-run]')
		if (survivor === null) throw new Error('the document has no runs left')
		const address = survivor.getAttribute('data-d2p-run')
		survivor.remove()

		const second = readSurface(dom, SAMPLE_IR)
		expect(second.anomalies.join('\n')).toContain(`run ${address} was removed`)
		expect(reconcile(SAMPLE_IR, second).slides[0]?.lane).toBe('drifted')
	})

	it('refuses a property the surface does not define instead of writing it', async () => {
		const dom = await documentOf(SAMPLE_IR)
		const span = dom.querySelector('[data-d2p-run]')
		if (span === null) throw new Error('the rendered document has no editable runs')
		span.setAttribute('data-d2p-props', JSON.stringify({ bold: true, fontFace: 'Comic Sans MS' }))

		const reading = readSurface(dom, SAMPLE_IR)
		expect(reading.anomalies.join('\n')).toContain('"fontFace"')
		expect(JSON.stringify(reconcile(SAMPLE_IR, reading).ir)).not.toContain('Comic Sans MS')
	})

	it('ignores a run address the model does not have', async () => {
		const dom = await documentOf(SAMPLE_IR)
		const span = dom.querySelector('[data-d2p-run]')
		if (span === null) throw new Error('the rendered document has no editable runs')
		const invented = span.cloneNode(true) as Element
		invented.setAttribute('data-d2p-run', 's1.sp999/0/0')
		span.parentElement?.append(invented)

		const reading = readSurface(dom, SAMPLE_IR)
		expect(reading.anomalies.join('\n')).toContain('s1.sp999/0/0 is in the document but not in the model')
	})
})
