import { describe, expect, it } from 'vitest'
import type { SlideModel } from '../../src/index'
import { convertDeck, convertSlide } from '../../src/index'

// Public-surface smoke test. `convertDeck` and `convertSlide` are both wired to
// the real (browser-only) engine and its extract path, so they need a real DOM:
// their behaviour is exercised in the
// browser/e2e layer (`test/browser/`), not here — this file only checks the
// exported surface and the IR type.

describe('dom2pptx public API', () => {
	it('exports convertDeck and convertSlide as functions', () => {
		expect(typeof convertDeck).toBe('function')
		expect(typeof convertSlide).toBe('function')
	})

	it('the IR slide model type is importable and structurally usable', () => {
		const model: SlideModel = {
			background: { type: 'color', value: 'FFFFFF' },
			items: [{ type: 'shape', position: { x: 0, y: 0, w: 1, h: 1 }, z: 0, fill: '250F6B' }],
		}
		expect(model.items).toHaveLength(1)
		expect(model.items[0].type).toBe('shape')
	})
})
