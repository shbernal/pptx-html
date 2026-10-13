/**
 * Writer round-trip, headless. The emit layer for a DOM-free model (colour
 * background plus one item of each kind the extractor emits: shape, line, text,
 * table, a bulleted list and a data-URL image) drives `pptx-ts`, the deck is
 * serialised and parsed back with the writer's `read` model to assert the
 * conversion round-trips. Browser `convertDeck` delivery is exercised by the
 * Playwright harness.
 */

import { ShapeType, TsPptx } from 'pptx-ts'
import { Presentation } from 'pptx-ts/read'
import { describe, expect, it } from 'vitest'
import type { SlideModel } from '../../src/heuristic/model'
import { addModelToSlide } from '../../src/heuristic/slide'
import { first } from '../support'

const SIZE = { width: 13.333, height: 7.5 }

// 1×1 transparent PNG — a minimal data-URL image the emit layer can place directly.
const PNG_1X1 =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR4nGNgAAIAAAUAAen63NgAAAAASUVORK5CYII='

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
			{ type: 'image', z: 5, position: { x: 7, y: 0.5, w: 2, h: 2 }, src: PNG_1X1 },
			{
				type: 'list',
				z: 6,
				position: { x: 7, y: 3, w: 5, h: 2 },
				text: [
					{ text: 'First priority', options: { bullet: {} } },
					{ text: 'Second priority', options: { bullet: {} } },
				],
				style: { fontFace: 'Aptos', fontSize: 18, color: '250F6B', align: 'left', valign: 'top' },
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
	it('emits every item kind onto ts-pptx without per-item issues', async () => {
		const { issues } = await emitToBytes(domFreeModel())
		expect(issues).toEqual([])
	})

	it('carries the heading and the bullet runs through to readable text', async () => {
		const { bytes } = await emitToBytes(domFreeModel())
		const pres = await Presentation.load(bytes)
		expect(pres.slides.length).toBe(1)
		const allText = first(pres.slides, 'slide')
			.shapes.map((shape: { text?: string }) => shape.text || '')
			.join(' ')
		expect(allText).toContain('Hello pptx-html')
		expect(allText).toContain('First priority')
		expect(allText).toContain('Second priority')
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
		// And no run is still on the writer's default. The model has a text box, a list and a
		// table, and only checking that *some* run moved would pass with the table
		// left behind — a table's cells do not take a table-level language. Matched on
		// `…Pr lang=` rather than on the bare attribute, which `altLang="en-US"`
		// contains as a substring on every run the writer emits.
		expect(slideXml).not.toMatch(/Pr lang="en-US"/)
	})
})
