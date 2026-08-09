/**
 * Browser/e2e: the opt-in `vectorizeSvg` path end to end in real Chromium. The static resolver returns a square `<path>` icon, so
 * with vectorization on, each icon's SVG raster image becomes an editable
 * `custGeom` path in the IR — and survives emit → ts-pptx → read as custom geometry.
 *
 * Requires Playwright: the full extract path needs real layout + an SVG DOM
 * (`DOMParser`) that jsdom/happy-dom don't provide faithfully.
 */

// @ts-expect-error — ts-pptx read entry typed via package exports.
import { Presentation } from '@shbernal/ts-pptx/read'
import { describe, expect, it } from 'vitest'
import { parseDeckHtml } from '../../src/heuristic/parse'
import { convertDeck, convertSlide } from '../../src/index'
import iconRowHtml from '../fixtures/icon-row.html?raw'
import type { RawModel } from './helpers'
import { base64ToBytes, staticResolveIcon } from './helpers'

async function modelFor(deckHtml: string, vectorizeSvg: boolean): Promise<RawModel> {
	const parsed = parseDeckHtml(deckHtml)
	const { model } = await convertSlide(parsed.headHTML, parsed.slides[0], {
		resolveIcon: staticResolveIcon,
		vectorizeSvg,
	})
	return model as unknown as RawModel
}

describe('vectorizeSvg — IR', () => {
	it('leaves SVG icons as raster images when off (no regression)', async () => {
		const model = await modelFor(iconRowHtml, false)
		expect(model.items.filter((i) => i.type === 'image').length).toBe(3)
		expect(model.items.filter((i) => i.type === 'path').length).toBe(0)
	})

	it('replaces each SVG icon image with an editable path when on', async () => {
		const model = await modelFor(iconRowHtml, true)
		const paths = model.items.filter((i) => i.type === 'path')
		const svgImages = model.items.filter((i) => i.type === 'image' && /^data:image\/svg/i.test((i.src as string) || ''))
		expect(svgImages.length).toBe(0)
		expect(paths.length).toBe(3)
		// The square icon path resolves to a concrete fill color (currentColor was
		// resolved by the extractor before vectorization).
		expect(paths.every((p) => typeof (p as { fill?: string }).fill === 'string')).toBe(true)
	})
})

describe('vectorizeSvg — emit → ts-pptx read', () => {
	it('exports the icons as readable custom-geometry shapes', async () => {
		const { base64 } = (await convertDeck(iconRowHtml, {
			resolveIcon: staticResolveIcon,
			vectorizeSvg: true,
			output: 'base64',
		})) as { base64: string }
		const pres = await Presentation.load(base64ToBytes(base64))
		const customs = pres.slides[0].shapes.filter((s: { customGeometry?: unknown }) => s.customGeometry)
		expect(customs.length).toBe(3)
		// Each square icon → moveTo + line segments + close.
		const cmds = customs[0].customGeometry.paths[0].commands.map((c: { cmd: string }) => c.cmd)
		expect(cmds[0]).toBe('moveTo')
		expect(cmds).toContain('close')
	})
})
