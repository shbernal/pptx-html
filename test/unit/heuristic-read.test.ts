/**
 * The boundary the `eval`'d extractor is quarantined behind.
 *
 * `heuristic/extractor.ts` is the one `@ts-nocheck` in the package, and the
 * argument for tolerating it is entirely this validation: the extractor is
 * unchecked *where it runs*, and everything downstream of it is checked because
 * `coerceSlideModel` refuses whatever does not fit `heuristic/model.ts`. So the
 * tests worth having here are about the refusals, not the happy path — a
 * boundary that accepts everything is a boundary in name only.
 *
 * They run in Node with no DOM on purpose: the validation is pure, and pinning
 * it here means the browser layer only has to prove the extractor produces
 * something, not that the checking works.
 */

import { describe, expect, it } from 'vitest'
import { coerceSlideModel } from '../../src/heuristic/read'

const RECT = { x: 1, y: 1, w: 2, h: 1 }

function slide(items: unknown[], rest: Record<string, unknown> = {}) {
	return { background: { type: 'color', value: 'FFFFFF' }, items, ...rest }
}

describe('the extractor boundary', () => {
	it('passes a well-formed model through with no complaints', () => {
		const read = coerceSlideModel(
			slide([
				{ type: 'text', z: 1, position: RECT, text: 'Hello' },
				{ type: 'shape', z: 2, position: RECT, fill: '451DC7' },
			])
		)
		expect(read.warnings).toEqual([])
		expect(read.model.items.map((item) => item.type)).toEqual(['text', 'shape'])
		expect(read.model.background).toEqual({ type: 'color', value: 'FFFFFF' })
	})

	it('throws when the extractor returned no model at all', () => {
		// A missing `items` array is a bug in this package, not a property of the
		// page. Continuing with an empty slide would report that bug as a blank
		// deck, which is the one outcome nobody can debug from the output.
		expect(() => coerceSlideModel(undefined)).toThrow(/no model/)
		expect(() => coerceSlideModel({ background: { type: 'color', value: 'FFF' } })).toThrow(/no model/)
	})

	it('drops an item with an unusable position and says which one', () => {
		const read = coerceSlideModel(
			slide([
				{ type: 'shape', z: 1, position: { x: 0, y: 0, w: Number.NaN, h: 1 }, label: 'hero panel' },
				{ type: 'shape', z: 2, position: RECT },
			])
		)
		expect(read.model.items).toHaveLength(1)
		expect(read.warnings).toEqual(['NEEDS REVIEW: "hero panel" had no usable position and was left out of the slide.'])
	})

	it('drops an item whose type it does not recognise', () => {
		const read = coerceSlideModel(slide([{ type: 'hologram', z: 1, position: RECT }]))
		expect(read.model.items).toEqual([])
		expect(read.warnings).toEqual(['NEEDS REVIEW: item 1 has an unrecognised type and was left out of the slide.'])
	})

	it('drops an item that carries no payload for its type', () => {
		// Each of these would reach the writer as a well-shaped item with nothing
		// in it, and surface as an exception naming neither the slide nor the
		// element. Named once here instead.
		const read = coerceSlideModel(
			slide([
				{ type: 'image', z: 1, position: RECT, src: '' },
				{ type: 'table', z: 2, position: RECT, rows: 'two' },
				{ type: 'path', z: 3, position: RECT, points: [] },
				{ type: 'text', z: 4, position: RECT },
			])
		)
		expect(read.model.items).toEqual([])
		expect(read.warnings).toEqual([
			'NEEDS REVIEW: a image has no source and was left out of the slide.',
			'NEEDS REVIEW: a table has no rows and was left out of the slide.',
			'NEEDS REVIEW: a path has no path points and was left out of the slide.',
			'NEEDS REVIEW: a text has no text and was left out of the slide.',
		])
	})

	it('keeps the option bags it cannot have a schema for', () => {
		// Runs, cell options and borders are opaque by design; validating them
		// would mean inventing a schema the writer does not publish, and would
		// throw away exactly the styling the lane exists to carry.
		const read = coerceSlideModel(
			slide([
				{
					type: 'text',
					z: 1,
					position: RECT,
					text: [{ text: 'Hi', options: { bold: true, hyperlink: { url: 'https://example.invalid' } } }],
					style: { fontSize: 28, align: 'center' },
					tag: 'h1',
				},
			])
		)
		expect(read.warnings).toEqual([])
		expect(read.model.items[0]).toMatchObject({
			tag: 'h1',
			style: { fontSize: 28, align: 'center' },
			text: [{ text: 'Hi', options: { bold: true, hyperlink: { url: 'https://example.invalid' } } }],
		})
	})

	it('falls back to a white background rather than failing the slide', () => {
		const read = coerceSlideModel(slide([], { background: { type: 'gradient', stops: [] } }))
		expect(read.model.background).toEqual({ type: 'color', value: 'FFFFFF' })
		expect(read.warnings).toEqual(['NEEDS REVIEW: the slide background could not be read and was left white.'])
	})

	it('carries an image background with its fallback colour', () => {
		const read = coerceSlideModel(slide([], { background: { type: 'image', src: 'x.png', fallback: '101010' } }))
		expect(read.model.background).toEqual({ type: 'image', src: 'x.png', fallback: '101010' })
	})

	it('defaults a missing z to 0 so paint order stays total', () => {
		const read = coerceSlideModel(slide([{ type: 'shape', position: RECT }]))
		expect(read.model.items[0].z).toBe(0)
	})

	it('keeps notes only when there are some', () => {
		expect(coerceSlideModel(slide([], { notes: 'Mention the oracle.' })).model.notes).toBe('Mention the oracle.')
		expect(coerceSlideModel(slide([], { notes: '' })).model).not.toHaveProperty('notes')
	})
})
