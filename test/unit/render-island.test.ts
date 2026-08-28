/**
 * The island's own rules, on a fixture rather than a corpus deck.
 *
 * These are the properties the return path is entitled to assume, and every one of them
 * fails *silently* if it is wrong: a truncated island parses as a smaller deck,
 * a hash over the wrong bytes reports tampering on an untouched document, and a
 * `surfaceHash` that moves with the whole model turns every text edit into
 * corruption. None of them would be caught by "does it render".
 */

import { describe, expect, it } from 'vitest'
import { renderDeck } from '../../src/render/document'
import { escapeForScript, INTEGRITY_ID, ISLAND_ID, integrityOf, islandTextOf } from '../../src/render/island'
import { SAMPLE_IR } from '../fixtures/render-ir'

const NO_ASSETS = { assets: 'ref' } as const

/** The island's text, taken the way the parser takes it: by block id. */
function islandOf(html: string, id: string = ISLAND_ID): string {
	const match = new RegExp(`<script type="application/json" id="${id}">(.*?)</script>`, 's').exec(html)
	if (match?.[1] === undefined) throw new Error(`no block with id ${id}`)
	return match[1]
}

describe('escaping', () => {
	it('keeps a literal </script> in the deck from ending the island', async () => {
		// The failure this prevents is not a mangled string: the HTML tokenizer would
		// end the script element early, the rest of the model would be parsed as
		// markup, and the document would still *look* fine.
		const hostile = structuredClone(SAMPLE_IR)
		const shape = hostile.slides[0]?.nodes.find((node) => node.kind === 'shape')
		if (shape?.kind !== 'shape' || !shape.text?.paragraphs[0]?.runs[0]) throw new Error('fixture shape changed')
		shape.text.paragraphs[0].runs[0].text = 'a </script><img src=x> b'

		const { html } = await renderDeck(hostile, NO_ASSETS)
		expect(html).not.toContain('a </script><img')
		expect(JSON.parse(islandOf(html))).toStrictEqual(hostile)
	})

	it('escapes only the character, not the value', () => {
		const json = JSON.stringify({ text: '</script> <!-- 5 < 6' })
		expect(escapeForScript(json)).not.toContain('<')
		expect(JSON.parse(escapeForScript(json))).toStrictEqual(JSON.parse(json))
	})
})

describe('the two hashes', () => {
	it('moves surfaceHash but not modelHash when a run is edited in place', async () => {
		// The distinction the whole scheme rests on. An edit inside the surface is a
		// *different model*, so `modelHash` moves too — what must hold is the
		// converse: something outside the surface must move `modelHash` while
		// leaving `surfaceHash` alone, so the parser can tell drift from an edit.
		const before = await integrityOf(SAMPLE_IR, islandTextOf(SAMPLE_IR))

		const moved = structuredClone(SAMPLE_IR)
		const shape = moved.slides[0]?.nodes.find((node) => node.kind === 'shape')
		if (shape?.kind !== 'shape' || shape.placement === null) throw new Error('fixture shape changed')
		shape.placement.box.x += 1

		const after = await integrityOf(moved, islandTextOf(moved))
		expect(after.modelHash).not.toBe(before.modelHash)
		expect(after.surfaceHash).toBe(before.surfaceHash)
	})

	it('moves both when the edit is inside the surface', async () => {
		const before = await integrityOf(SAMPLE_IR, islandTextOf(SAMPLE_IR))

		const edited = structuredClone(SAMPLE_IR)
		const shape = edited.slides[0]?.nodes.find((node) => node.kind === 'shape')
		if (shape?.kind !== 'shape' || !shape.text?.paragraphs[0]?.runs[0]) throw new Error('fixture shape changed')
		shape.text.paragraphs[0].runs[0].text += ' edited'

		const after = await integrityOf(edited, islandTextOf(edited))
		expect(after.surfaceHash).not.toBe(before.surfaceHash)
	})

	it('hashes the bytes it embeds', async () => {
		// If these ever disagree, every untouched document reports as tampered with.
		const { html, integrity } = await renderDeck(SAMPLE_IR, NO_ASSETS)
		const embedded = islandOf(html)
		expect(embedded).toBe(islandTextOf(SAMPLE_IR))

		const stated = JSON.parse(islandOf(html, INTEGRITY_ID))
		expect(stated).toStrictEqual(integrity)
	})
})

describe('the renderer is read-only', () => {
	it('leaves the IR it was given untouched', async () => {
		// The trap the renderer is most likely to fall into later: normalizing a colour or
		// clamping a box back into the model. It breaks Invariant R somewhere else
		// entirely, so it is asserted here rather than left to review.
		const untouched = structuredClone(SAMPLE_IR)
		await renderDeck(SAMPLE_IR, NO_ASSETS)
		expect(SAMPLE_IR).toStrictEqual(untouched)
	})
})

describe('line ends', () => {
	it('draws modeled arrowheads as stroke-relative SVG markers', async () => {
		// The fixture connector carries a medium triangle tail. Its marker belongs in
		// the slide-local defs and the path names it; neither arm may retain the old
		// approximation marker once the line end is actually drawn.
		const { html } = await renderDeck(SAMPLE_IR, NO_ASSETS)
		expect(html).toContain('markerUnits="strokeWidth"')
		expect(html).toContain('orient="auto-start-reverse"')
		expect(html).toMatch(/marker-end="url\(#pxh-p\d+\)"/)
		expect(html).not.toContain('line:ends')
	})

	it('has distinct marker geometry for every modeled line-end type', async () => {
		const expected = {
			triangle: 'M 0 0 L 10 5 L 0 10 Z',
			stealth: 'M 0 0 L 10 5 L 0 10 L 3 5 Z',
			diamond: 'M 0 5 L 5 0 L 10 5 L 5 10 Z',
			oval: '<ellipse cx="5" cy="5" rx="5" ry="4"',
			arrow: 'M 0 0 L 10 5 L 0 10',
		} as const
		for (const [type, geometry] of Object.entries(expected)) {
			const ir = structuredClone(SAMPLE_IR)
			const connector = ir.slides[0]?.nodes.find((node) => node.kind === 'connector')
			if (connector?.kind !== 'connector' || connector.stroke.kind !== 'line')
				throw new Error('fixture connector changed')
			connector.stroke.tail = { type: type as keyof typeof expected, width: 'lg', length: 'sm' }
			const { html } = await renderDeck(ir, NO_ASSETS)
			expect([type, html.includes(geometry)]).toStrictEqual([type, true])
			expect(html).toContain('markerWidth="1.5" markerHeight="3.75"')
		}
	})
})
