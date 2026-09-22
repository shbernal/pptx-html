/**
 * Writer round-trip, headless. The emit layer for a DOM-free model (colour
 * background + shape + line + text + table) drives `pptx-ts`, the deck
 * is serialised and parsed back with the writer's `read` model to assert the
 * conversion round-trips. Browser `convertDeck`
 * delivery is exercised by the Playwright harness.
 */

import { ShapeType, TsPptx } from 'pptx-ts'
import { Presentation } from 'pptx-ts/read'
import { describe, expect, it } from 'vitest'
import type { SlideModel } from '../../src/heuristic/model'
import { addModelToSlide } from '../../src/heuristic/slide'
import { first } from '../support'

const SIZE = { width: 13.333, height: 7.5 }

// A model in the *runtime* shape the emit layer reads (text/table items carry a
// `style` object and `{ text, options }` runs), none of whose items needs a DOM.
function domFreeModel(): SlideModel {
	return {
		background: { type: 'color', value: '1A1A2E' },
		items: [
			{ type: 'shape', z: 1, position: { x: 0.5, y: 0.5, w: 3, h: 1.2 }, fill: '451DC7', radius: 0.1 },
			{ type: 'line', z: 2, position: { x: 0.5, y: 2, w: 4, h: 0 }, color: '04F06A', width: 2, dash: 'solid' },
			{
				type: 'text',
				z: 3,
				tag: 'h1',
				position: { x: 0.5, y: 2.4, w: 6, h: 0.8 },
				text: [{ text: 'Hello pptx-html', options: {} }],
				style: { fontFace: 'Aptos', fontSize: 28, color: 'FFFFFF', bold: true, align: 'left', valign: 'top' },
			},
			{
				type: 'table',
				z: 4,
				position: { x: 0.5, y: 3.6, w: 6, h: 1.5 },
				rows: [
					[
						{ text: 'A', options: {} },
						{ text: 'B', options: {} },
					],
					[
						{ text: 'C', options: {} },
						{ text: 'D', options: {} },
					],
				],
			},
		],
	}
}

// `toBytes()` returns exactly what `Presentation.load` reads, so the deck goes
// from writer to reader without a base64 detour.
async function emitToBytes(model: SlideModel) {
	const pptx = new TsPptx()
	pptx.layout = 'LAYOUT_16x9'
	const slide = pptx.addSlide()
	const issues = await addModelToSlide({ ShapeType }, slide, model, SIZE, 'en')
	const bytes = await pptx.toBytes()
	return { issues, bytes }
}

describe('emit → ts-pptx → read round-trip', () => {
	it('emits a DOM-free model onto ts-pptx without per-item issues', async () => {
		const { issues } = await emitToBytes(domFreeModel())
		expect(issues).toEqual([])
	})

	it('produces a single readable slide with shapes', async () => {
		const { bytes } = await emitToBytes(domFreeModel())
		const pres = await Presentation.load(bytes)
		expect(pres.slides.length).toBe(1)
		expect(first(pres.slides, 'slide').shapes.length).toBeGreaterThan(0)
	})

	it('carries the heading text through to the OOXML', async () => {
		const { bytes } = await emitToBytes(domFreeModel())
		const pres = await Presentation.load(bytes)
		const allText = first(pres.slides, 'slide')
			.shapes.map((shape: { text?: string }) => shape.text || '')
			.join(' ')
		expect(allText).toContain('Hello pptx-html')
	})

	it('puts the document language on the runs, which is the only place it fits', async () => {
		// The lane parses `<html lang>` and used to assign it to `pptx.lang` and to
		// `theme.lang`, neither of which the writer has: both landed as own properties
		// on the instance and nothing wrote them into the package. `lang` is a *run*
		// option, so this asserts on the slide XML — the read model exposes no run
		// language to check instead, and "the assignment happened" is exactly the
		// thing that was true before and still meant nothing.
		const pptx = new TsPptx()
		pptx.layout = 'LAYOUT_16x9'
		const slide = pptx.addSlide()
		await addModelToSlide({ ShapeType }, slide, domFreeModel(), SIZE, 'fr-CA')
		const pres = await Presentation.load(await pptx.toBytes())
		const slideXml = new TextDecoder().decode(first(pres.slides, 'slide').part.serialize())
		expect(slideXml).toContain('lang="fr-CA"')
		// And no run is still on the writer's default. The model has a text box and a
		// table, and only checking that *some* run moved would pass with the table
		// left behind — a table's cells do not take a table-level language. Matched on
		// `…Pr lang=` rather than on the bare attribute, which `altLang="en-US"`
		// contains as a substring on every run the writer emits.
		expect(slideXml).not.toMatch(/Pr lang="en-US"/)
	})

	it('states the slide size the layout asked for', async () => {
		// This used to assert `type="screen16x9"`, which a post-write repair stamped
		// on unconditionally — including onto decks that were not 16:9. It went
		// unnoticed because this deck *is* 16:9: the writer's `LAYOUT_16x9` is 10in ×
		// 5.625in, so the wrong rule and the right answer agreed here. The attribute
		// is optional and PowerPoint infers it from the dimensions, which are the
		// fact worth pinning.
		const { bytes } = await emitToBytes(domFreeModel())
		const pres = await Presentation.load(bytes)
		const presXml = new TextDecoder().decode(pres.presentationPart.serialize())
		expect(presXml).toMatch(/<p:sldSz[^>]*cx="9144000"[^>]*cy="5143500"/)
	})
})
