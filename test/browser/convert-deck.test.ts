/**
 * Browser/e2e layer. Runs a full multi-slide deck through
 * `convertDeck` end-to-end in real Chromium, exercising the `opts.output`
 * delivery modes that replaced the old `window.__TEST__` hack, and parses the
 * `base64` result back with ts-pptx's `read` model to assert deck structure.
 */

import { describe, expect, it } from 'vitest'
import { convertDeck } from '../../src/index'
import deckHtml from '../fixtures/deck.html?raw'
import { allDeckText, loadDeck, staticResolveIcon } from './helpers'

describe('convertDeck delivery modes', () => {
	it('output: base64 → two readable slides carrying their text', async () => {
		const result = await convertDeck(deckHtml, { output: 'base64', resolveIcon: staticResolveIcon })
		expect(result.slideCount).toBe(2)
		expect(result.totalSlideCount).toBe(2)
		expect(result.stopped).toBe(false)
		expect(typeof result.base64).toBe('string')

		const deck = await loadDeck(result.base64 as string)
		expect(deck.slides.length).toBe(2)
		const text = allDeckText(deck)
		expect(text).toContain('FY26 Strategy')
		expect(text).toContain('Key Priorities')
		expect(text).toContain('Accelerate cloud migration')
	})

	it('output: blob → a non-empty pptx blob', async () => {
		const result = await convertDeck(deckHtml, { output: 'blob', resolveIcon: staticResolveIcon })
		expect(result.base64).toBeUndefined()
		expect(result.blob).toBeInstanceOf(Blob)
		expect((result.blob as Blob).size).toBeGreaterThan(0)
	})

	it('output: pptx-instance → the live writer (pre-repair)', async () => {
		const result = await convertDeck(deckHtml, { output: 'pptx-instance', resolveIcon: staticResolveIcon })
		expect(result.base64).toBeUndefined()
		expect(result.pptx).toBeTruthy()
		expect(typeof (result.pptx as { write?: unknown }).write).toBe('function')
	})

	it('reports cooperative stop via shouldStop', async () => {
		// shouldStop is polled before and after each slide; allow the first slide
		// through, then request a stop. The deck stops with one slide created.
		let calls = 0
		const result = await convertDeck(deckHtml, {
			output: 'base64',
			resolveIcon: staticResolveIcon,
			shouldStop: () => calls++ >= 1,
		})
		expect(result.stopped).toBe(true)
		expect(result.slideCount).toBe(1)
		expect(result.totalSlideCount).toBe(2)
	})
})
