/**
 * Headless emit coverage. Complements `emit-fork.test.ts`
 * (shape/line/text/table) by exercising the remaining IR item types — `list`
 * (bulleted runs) and `image` (a data-URL picture) — through the emit layer on
 * ts-pptx, then parsing the deck back with ts-pptx's `read` model. No DOM
 * needed: these item shapes are hand-written, exactly as `heuristic/extractor.ts` would emit.
 */

import { ShapeType, TsPptx } from '@shbernal/ts-pptx'
import { Presentation } from '@shbernal/ts-pptx/read'
import { describe, expect, it } from 'vitest'
import type { SlideModel } from '../../src/heuristic/model'
import { addModelToSlide } from '../../src/heuristic/slide'
import { first } from '../support'

const SIZE = { width: 13.333, height: 7.5 }

// 1×1 transparent PNG — a minimal data-URL image the emit layer can place directly.
const PNG_1X1 =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR4nGNgAAIAAAUAAen63NgAAAAASUVORK5CYII='

function listAndImageModel(): SlideModel {
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

// `toBytes()` returns exactly what `Presentation.load` reads, so the deck goes
// from writer to reader without a base64 detour.
async function emitToBytes(model: SlideModel) {
	const pptx = new TsPptx()
	pptx.layout = 'LAYOUT_16x9'
	const slide = pptx.addSlide()
	const issues = await addModelToSlide({ ShapeType }, slide, model, SIZE)
	const bytes = await pptx.toBytes()
	return { issues, bytes }
}

describe('emit list + image → ts-pptx → read round-trip', () => {
	it('emits list and image items without per-item issues', async () => {
		const { issues } = await emitToBytes(listAndImageModel())
		expect(issues).toEqual([])
	})

	it('carries the bullet runs through to readable text', async () => {
		const { bytes } = await emitToBytes(listAndImageModel())
		const pres = await Presentation.load(bytes)
		expect(pres.slides.length).toBe(1)
		const allText = first(pres.slides, 'slide')
			.shapes.map((shape: { text?: string }) => shape.text || '')
			.join(' ')
		expect(allText).toContain('First priority')
		expect(allText).toContain('Second priority')
	})
})
