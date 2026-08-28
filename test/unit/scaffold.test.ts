import { describe, expect, it } from 'vitest'
import type { Background, heuristic } from '../../src/index'
import { convertDeck, convertSlide, emitDeck, importDeck, parseDeck, project, renderDeck } from '../../src/index'
import { first } from '../support'

// Public-surface smoke test. What it is actually pinning is the *shape* of the
// surface: four legs of one loop, plus a second entry point that infers, and two
// model vocabularies that must stay distinguishable at the top level.
//
// Behaviour lives elsewhere — the loop in `test/oracle/`, the browser-only
// heuristic lane in `test/browser/`. This file checks only that the names are
// there and that the two models do not quietly become one.

describe('pptx-html public API', () => {
	it('exports the four legs of the loop', () => {
		expect([importDeck, renderDeck, parseDeck, emitDeck].map((leg) => typeof leg)).toEqual([
			'function',
			'function',
			'function',
			'function',
		])
	})

	it('exports the heuristic lane and the editable surface beside it', () => {
		expect(typeof convertDeck).toBe('function')
		expect(typeof convertSlide).toBe('function')
		expect(typeof project).toBe('function')
	})

	it('keeps the two models apart at the type level', () => {
		// `Background` means different things in the two lanes — a paint-model
		// union in `RenderIr`, a DOM-shaped `{ type, value }` in the heuristic
		// lane's model. Before the namespace, `export type *` from both would have
		// made one of them silently win. This block only compiles while they are
		// separately nameable.
		const inferred: heuristic.SlideModel = {
			background: { type: 'color', value: 'FFFFFF' },
			items: [{ type: 'shape', position: { x: 0, y: 0, w: 1, h: 1 }, z: 0, fill: '250F6B' }],
		}
		const modeled: Background = { source: 'master', fill: { kind: 'solid', color: { kind: 'srgb', hex: 'FFFFFF' } } }

		expect(first(inferred.items, 'item').type).toBe('shape')
		expect(modeled.source).toBe('master')
	})
})
