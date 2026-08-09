/**
 * Browser/extract layer. Loads representative slide fixtures, runs them through
 * the full DOM → IR path (`convertSlide`, which
 * renders in a hidden frame and reads computed styles), and asserts the IR
 * structure. Requires real Chromium (Playwright): jsdom/happy-dom lack the
 * `getBoundingClientRect` / `getComputedStyle` layout and canvas these rely on.
 *
 * Snapshots are of a *position-free* summary (see `summarizeModel`): pixel
 * positions depend on font metrics, but the structural shape is stable.
 */

import { describe, expect, it } from 'vitest'
import { parseDeckHtml } from '../../src/heuristic/parse'
import { convertSlide } from '../../src/index'
import bulletsHtml from '../fixtures/bullets.html?raw'
import coverHtml from '../fixtures/cover.html?raw'
import iconRowHtml from '../fixtures/icon-row.html?raw'
import tableHtml from '../fixtures/table.html?raw'
import type { RawModel } from './helpers'
import { runText, staticResolveIcon, summarizeModel } from './helpers'

async function modelFor(deckHtml: string): Promise<RawModel> {
	const parsed = parseDeckHtml(deckHtml)
	const { model } = await convertSlide(parsed.headHTML, parsed.slides[0], { resolveIcon: staticResolveIcon })
	return model as unknown as RawModel
}

const textItems = (model: RawModel) => model.items.filter((item) => item.type === 'text')

describe('cover slide', () => {
	it('extracts a heading, subtitle and a gradient background image', async () => {
		const model = await modelFor(coverHtml)
		expect((model.background as { type: string }).type).toBe('color')

		const heading = textItems(model).find((item) => runText(item.text).includes('Quarterly Strategy Review'))
		expect(heading).toBeTruthy()
		expect(heading?.tag).toBe('h1')
		expect(heading?.style?.bold).toBe(true)
		expect(heading?.style?.color).toBe('FFFFFF')

		const subtitle = textItems(model).find((item) => runText(item.text).includes('Northwind'))
		expect(subtitle?.style?.color).toBe('04F06A')

		// The CSS gradient becomes a rasterised, full-bleed decor image behind everything.
		const gradient = model.items.find((item) => item.type === 'image')
		expect(gradient).toBeTruthy()
		expect(gradient?.z).toBeLessThan(0)

		expect(summarizeModel(model)).toMatchSnapshot()
	})
})

describe('bullets slide', () => {
	it('extracts a heading and a list with one run per bullet', async () => {
		const model = await modelFor(bulletsHtml)

		const heading = textItems(model).find((item) => runText(item.text).includes('Key Priorities'))
		expect(heading?.tag).toBe('h2')

		const list = model.items.find((item) => item.type === 'list')
		expect(list).toBeTruthy()
		const runs = Array.isArray(list?.text) ? list.text.map((run) => run.text) : []
		expect(runs).toEqual(['Accelerate cloud migration', 'Strengthen data governance', 'Expand the partner ecosystem'])

		expect(summarizeModel(model)).toMatchSnapshot()
	})
})

describe('table slide', () => {
	it('extracts a 3×3 table with header and body cells', async () => {
		const model = await modelFor(tableHtml)

		const table = model.items.find((item) => item.type === 'table')
		expect(table).toBeTruthy()
		const rows = (table?.rows || []).map((row) => row.map((cell) => cell.text))
		expect(rows.length).toBe(3)
		expect(rows[0]).toEqual(['Workstream', 'Owner', 'Status'])
		expect(rows[1]).toEqual(['Cloud migration', 'A. Dupont', 'On track'])
		expect(rows[2][2]).toBe('At risk')

		expect(summarizeModel(model)).toMatchSnapshot()
	})
})

describe('icon row slide', () => {
	it('inlines the injected resolver SVG into one image per icon', async () => {
		const model = await modelFor(iconRowHtml)

		const images = model.items.filter((item) => item.type === 'image')
		expect(images.length).toBe(3)
		// Every icon resolved through the injected static resolver → inline SVG image.
		expect(images.every((item) => /^data:image\/svg/i.test(item.src || ''))).toBe(true)

		expect(summarizeModel(model)).toMatchSnapshot()
	})
})
