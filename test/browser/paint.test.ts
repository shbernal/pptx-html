/**
 * Whether the picture is legible — the one claim about the visual channel that
 * needs a browser doing layout to answer.
 *
 * Everything else about the renderer is checked by reading its output as a
 * string, and that is usually the right level: the island is what the round trip
 * trusts, and `render.test.ts` proves the island survives whatever the picture
 * does. The gap this file closes is that a *string* assertion cannot see a
 * browser refuse a length. The renderer once wrote text sizes in EMU-as-px
 * (`font-size: 558800px` for a 44pt run, correct arithmetic in a viewBox measured
 * in EMU) and browsers cap `font-size` — Chrome at 10000px. Every run above about
 * a point painted at the same hairline height, so every slide previewed blank.
 *
 * Not one test went red. The unit layer, the oracle and the surface tests all
 * passed, because all three read the model and none of them looked at the paint.
 * So the assertion here is deliberately about painted geometry rather than about
 * the markup that produced it: it is only worth anything if it would fail for a
 * length the browser silently declines to honour.
 */

import { describe, expect, it } from 'vitest'
import type { RenderIr } from '../../src/ir/render'
import { renderDeck } from '../../src/render/document'
import { EMU_PER_POINT } from '../../src/render/paint'
import { SAMPLE_IR } from '../fixtures/render-ir'
import { at } from '../support'

const NO_ASSETS = { assets: 'ref' } as const

/** A real iframe, because `DOMParser` gives a DOM with no layout and no styles. */
async function paint(ir: RenderIr): Promise<{ doc: Document; dispose: () => void }> {
	const { html } = await renderDeck(ir, NO_ASSETS)
	const frame = document.createElement('iframe')
	frame.style.cssText = 'position:fixed;left:-9999px;top:0;width:1280px;height:720px;border:0'
	document.body.append(frame)
	await new Promise<void>((resolve) => {
		frame.addEventListener('load', () => resolve(), { once: true })
		frame.srcdoc = html
	})
	const doc = frame.contentDocument
	if (doc === null) throw new Error('the preview frame produced no document')
	return { doc, dispose: () => frame.remove() }
}

/** CSS pixels per EMU, measured from the slide as it was actually laid out. */
function scaleOf(doc: Document, ir: RenderIr): number {
	const svg = doc.querySelector('svg')
	if (svg === null) throw new Error('the rendered document has no slide')
	return svg.getBoundingClientRect().width / ir.size.w
}

describe('the painted slide', () => {
	it('paints a run at the point size it states', async () => {
		const { doc, dispose } = await paint(SAMPLE_IR)
		try {
			// `Quarterly ` states 40pt. Its line box is that size plus leading, so the
			// band is wide — the failure being guarded against is off by a factor of
			// thirty, not by a few percent.
			const span = doc.querySelector('[data-pxh-run="s1.sp2/0/0"]')
			if (span === null) throw new Error('the rendered document has no first run')
			const points = span.getBoundingClientRect().height / scaleOf(doc, SAMPLE_IR) / EMU_PER_POINT

			expect(points).toBeGreaterThan(30)
			expect(points).toBeLessThan(70)
		} finally {
			dispose()
		}
	})

	it('asks for no font size the browser declines to honour', async () => {
		// The general form of the bug, and the one that would catch it again in a
		// different unit. Comparing the size the renderer wrote against the size the
		// browser computed says nothing about typography and everything about whether
		// the length was accepted: a capped value is silently substituted, so the two
		// agree exactly until they do not agree at all.
		const { doc, dispose } = await paint(SAMPLE_IR)
		try {
			const compared: number[] = []
			for (const span of doc.querySelectorAll('[data-pxh-run]')) {
				const declared = /font-size:([\d.]+)px/.exec(span.getAttribute('style') ?? '')
				if (declared === null) continue
				const computed = Number.parseFloat(doc.defaultView?.getComputedStyle(span).fontSize ?? '')
				expect(computed).toBeCloseTo(Number.parseFloat(at(declared, 1, 'font-size capture')), 3)
				compared.push(computed)
			}
			// Otherwise a renderer that stopped stating sizes altogether would pass by
			// comparing nothing.
			expect(compared.length).toBeGreaterThan(2)
		} finally {
			dispose()
		}
	})

	it('paints text inside the box that holds it', async () => {
		// Catches the other direction: a unit slip that made text enormous would sail
		// past a "bigger than a hairline" assertion.
		const { doc, dispose } = await paint(SAMPLE_IR)
		try {
			const span = doc.querySelector('[data-pxh-run="s1.sp2/0/0"]')
			const frame = doc.querySelector('[data-pxh-node="s1.sp2"] foreignObject')
			if (span === null || frame === null) throw new Error('the rendered document lost its title')
			const text = span.getBoundingClientRect()
			const box = frame.getBoundingClientRect()
			expect(text.width).toBeLessThanOrEqual(box.width + 1)
			expect(text.height).toBeLessThanOrEqual(box.height + 1)
		} finally {
			dispose()
		}
	})
})
