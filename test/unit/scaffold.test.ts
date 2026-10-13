import { describe, expect, it } from 'vitest'
import type { Background, heuristic } from '../../src/index'
import { first } from '../support'

// The two model vocabularies must stay separately nameable at the top level of
// the public surface. That the names exist at all is already checked twice over:
// `tsc` fails on an import of a missing export, and `examples/round-trip.mjs` and
// `site/` import the same names through `dist` at run time.

describe('pptx-html public API', () => {
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
