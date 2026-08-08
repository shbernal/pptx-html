/**
 * Headless emit coverage. Complements `emit-fork.test.ts`
 * (shape/line/text/table) by exercising the remaining IR item types — `list`
 * (bulleted runs) and `image` (a data-URL picture) — through the emit layer on
 * ts-pptx, then parsing the deck back with ts-pptx's `read` model. No DOM
 * needed: these item shapes are hand-written, exactly as `extract/` would emit.
 */

// @ts-expect-error — ts-pptx ships its own types; node-resolved entry is fine for tests.
import { ShapeType, TsPptx } from '@shbernal/ts-pptx'
// @ts-expect-error — read entry typed via package exports.
import { Presentation } from '@shbernal/ts-pptx/read'
import { describe, expect, it } from 'vitest'
import { addModelToSlide } from '../../src/emit/slide'
import { repairPptxBase64 } from '../../src/repair/repair'

const SIZE = { width: 13.333, height: 7.5 }

// 1×1 transparent PNG — a minimal data-URL image the emit layer can place directly.
const PNG_1X1 =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR4nGNgAAIAAAUAAen63NgAAAAASUVORK5CYII='

function listAndImageModel() {
	return {
		background: { type: 'color', value: 'FFFFFF' },
		items: [
			{ type: 'image', z: 1, position: { x: 1, y: 1, w: 2, h: 2 }, src: PNG_1X1 },
			{
				type: 'list',
				z: 2,
				position: { x: 1, y: 4, w: 6, h: 2 },
				text: [
					{ text: 'First priority', options: { bullet: {} } },
					{ text: 'Second priority', options: { bullet: {} } },
				],
				style: { fontFace: 'Aptos', fontSize: 18, color: '250F6B', align: 'left', valign: 'top' },
			},
		],
	}
}

async function emitToBase64(model: ReturnType<typeof listAndImageModel>) {
	const pptx = new TsPptx()
	pptx.layout = 'LAYOUT_16x9'
	const slide = pptx.addSlide()
	const issues = await addModelToSlide({ ShapeType }, slide, model, SIZE)
	const base64 = await repairPptxBase64(await pptx.write({ outputType: 'base64' }))
	return { issues, base64 }
}

function loadDeck(base64: string) {
	return Presentation.load(Uint8Array.from(Buffer.from(base64, 'base64')))
}

describe('emit list + image → ts-pptx → read round-trip', () => {
	it('emits list and image items without per-item issues', async () => {
		const { issues } = await emitToBase64(listAndImageModel())
		expect(issues).toEqual([])
	})

	it('carries the bullet runs through to readable text', async () => {
		const { base64 } = await emitToBase64(listAndImageModel())
		const pres = await loadDeck(base64)
		expect(pres.slides.length).toBe(1)
		const allText = pres.slides[0].shapes.map((shape: { text?: string }) => shape.text || '').join(' ')
		expect(allText).toContain('First priority')
		expect(allText).toContain('Second priority')
	})
})
