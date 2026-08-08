/**
 * Writer round-trip, headless. The emit layer for a DOM-free model (colour
 * background + shape + line + text + table) drives `@shbernal/ts-pptx`, the deck
 * is serialised, run through the repair pass, and parsed back with the writer's
 * `read` model to assert the conversion round-trips. Browser `convertDeck`
 * delivery is exercised by the Playwright harness.
 */

// @ts-expect-error — ts-pptx ships its own types; node-resolved entry is fine for tests.
import { ShapeType, TsPptx } from '@shbernal/ts-pptx'
// @ts-expect-error — read entry typed via package exports.
import { Presentation } from '@shbernal/ts-pptx/read'
import { describe, expect, it } from 'vitest'
import { addModelToSlide } from '../../src/emit/slide'
import { repairPptxBase64 } from '../../src/repair/repair'

const SIZE = { width: 13.333, height: 7.5 }

// A model in the *runtime* shape the emit layer reads (text/table items carry a
// `style` object and `{ text, options }` runs), none of whose items needs a DOM.
function domFreeModel() {
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
				text: [{ text: 'Hello dom2pptx', options: {} }],
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

async function emitToBase64(model: ReturnType<typeof domFreeModel>) {
	const pptx = new TsPptx()
	pptx.layout = 'LAYOUT_16x9'
	const slide = pptx.addSlide()
	const issues = await addModelToSlide({ ShapeType }, slide, model, SIZE)
	const base64 = await repairPptxBase64(await pptx.write({ outputType: 'base64' }))
	return { issues, base64 }
}

// ts-pptx's `Presentation.load` reads raw bytes; decode the base64 deck first.
function loadDeck(base64: string) {
	return Presentation.load(Uint8Array.from(Buffer.from(base64, 'base64')))
}

describe('emit → ts-pptx → repair → read round-trip', () => {
	it('emits a DOM-free model onto ts-pptx without per-item issues', async () => {
		const { issues } = await emitToBase64(domFreeModel())
		expect(issues).toEqual([])
	})

	it('produces a single readable slide with shapes', async () => {
		const { base64 } = await emitToBase64(domFreeModel())
		const pres = await loadDeck(base64)
		expect(pres.slides.length).toBe(1)
		expect(pres.slides[0].shapes.length).toBeGreaterThan(0)
	})

	it('carries the heading text through to the OOXML', async () => {
		const { base64 } = await emitToBase64(domFreeModel())
		const pres = await loadDeck(base64)
		const allText = pres.slides[0].shapes.map((shape: { text?: string }) => shape.text || '').join(' ')
		expect(allText).toContain('Hello dom2pptx')
	})

	it('repair pass types the slide size as screen16x9', async () => {
		const { base64 } = await emitToBase64(domFreeModel())
		const pres = await loadDeck(base64)
		const presXml = new TextDecoder().decode(pres.presentationPart.bytes)
		expect(presXml).toMatch(/type="screen16x9"/)
	})
})
