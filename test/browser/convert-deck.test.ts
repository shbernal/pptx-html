/**
 * Browser/e2e layer. Runs a full multi-slide deck through
 * `convertDeck` end-to-end in real Chromium, exercising the `opts.output`
 * delivery modes that replaced the old `window.__TEST__` hack, and parses the
 * `base64` result back with ts-pptx's `read` model to assert deck structure.
 */

import { TsPptx } from '@shbernal/ts-pptx'
import { Presentation } from '@shbernal/ts-pptx/read'
import { describe, expect, it } from 'vitest'
import { convertDeck } from '../../src/index'
import deckHtml from '../fixtures/deck.html?raw'
import { allDeckText, loadDeck, staticResolveIcon } from './helpers'

/**
 * A writer that does not answer to `toBytes()`, which is what `opts.pptxFactory`
 * is allowed to hand back: the method arrived in ts-pptx 3.3.0, and the seam
 * predates it. Forwarding rather than subclassing keeps `this` on the real
 * instance, so nothing but the one missing method differs.
 */
function writerWithoutToBytes() {
	const real = new TsPptx()
	return {
		get layout() {
			return real.layout
		},
		set layout(value: string) {
			real.layout = value
		},
		get presLayout() {
			return real.presLayout
		},
		set author(value: string) {
			real.author = value
		},
		set subject(value: string) {
			real.subject = value
		},
		set company(value: string) {
			real.company = value
		},
		set theme(value: Record<string, unknown>) {
			real.theme = value
		},
		// `lang` is deliberately absent: ts-pptx has no deck-level setter for it, so
		// the engine's assignment lands on this object exactly as it lands on a real
		// writer instance.
		addSlide: () => real.addSlide(),
		write: (opts: { outputType: 'base64' }) => real.write(opts),
	}
}

async function bytesOfBlob(blob: Blob): Promise<Uint8Array> {
	return new Uint8Array(await blob.arrayBuffer())
}

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

	it('output: blob → a readable pptx, straight from the writer bytes', async () => {
		const result = await convertDeck(deckHtml, { output: 'blob', resolveIcon: staticResolveIcon })
		expect(result.base64).toBeUndefined()
		expect(result.blob).toBeInstanceOf(Blob)
		// Read the deck back rather than checking the blob is non-empty. This path
		// takes `toBytes()` now, and "some bytes arrived" would pass just as well
		// for a blob holding the wrong thing.
		const deck = await Presentation.load(await bytesOfBlob(result.blob as Blob))
		expect(deck.slides.length).toBe(2)
	})

	it('output: blob → falls back to the base64 encoder for a writer with no toBytes()', async () => {
		const result = await convertDeck(deckHtml, {
			output: 'blob',
			resolveIcon: staticResolveIcon,
			pptxFactory: writerWithoutToBytes,
		})
		const deck = await Presentation.load(await bytesOfBlob(result.blob as Blob))
		expect(deck.slides.length).toBe(2)
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
