/**
 * Finding the island, and the escaping hazard that decided how it is hashed.
 *
 * The corpus lanes cannot reach any of this: they render decks the writer
 * produced, and none of those contains a backslash in a run of text or a stray
 * `id="dom2pptx-ir"` in its markup. Each case below is constructed, because each
 * one fails by producing a *plausible* wrong answer rather than an error.
 */

import { describe, expect, it } from 'vitest'
import type { RenderIr } from '../../src/ir/render'
import { parseDeck } from '../../src/parse/deck'
import { blocksOf } from '../../src/parse/island'
import { renderDeck } from '../../src/render/document'
import { ISLAND_ID, islandTextOf } from '../../src/render/island'
import { SAMPLE_IR } from '../fixtures/render-ir'

const NO_DOM = { parseHtml: null } as const
const NO_ASSETS = { assets: 'ref' } as const

/** `SAMPLE_IR` with one run's text replaced. */
function withText(text: string): RenderIr {
	const ir = structuredClone(SAMPLE_IR) as RenderIr
	const shape = ir.slides[0]?.nodes.find((node) => node.kind === 'shape')
	if (shape?.kind !== 'shape' || !shape.text?.paragraphs[0]?.runs[0]) throw new Error('fixture shape changed')
	shape.text.paragraphs[0].runs[0].text = text
	return ir
}

describe('the escape has no inverse, and the hash does not need one', () => {
	it('round-trips a run holding the literal characters of a unicode escape', () => {
		// The case that decided `modelHash` covers the *embedded* text rather than
		// the raw JSON. This run's text is a backslash followed by `u003c`, which
		// `JSON.stringify` writes as a doubled backslash — indistinguishable, to a
		// string replacement, from the escape the renderer applies to a real `<`.
		// A reader that "un-escaped" before hashing would corrupt exactly this model
		// and report it as tampered with.
		const ir = withText('\\u003c and \\\\u003c')
		const text = islandTextOf(ir)
		expect(JSON.parse(text)).toStrictEqual(ir)
	})

	it('verifies a document whose deck is full of angle brackets', async () => {
		const ir = withText('</script><!-- 5 < 6 & 7 > 6 -->')
		const { html } = await renderDeck(ir, NO_ASSETS)
		const parsed = await parseDeck(html, NO_DOM)
		expect(parsed.island).toStrictEqual(ir)
	})
})

describe('the block scanner', () => {
	it('does not mistake text that mentions the id for the block itself', () => {
		// Reading this as an island would turn "not one of our pages" into a parse
		// error on a page that is perfectly fine.
		const blocks = blocksOf(`<p>look for id="${ISLAND_ID}" in the source</p>`)
		expect(blocks.island).toBeNull()
	})

	it('finds the block when the attributes are in the other order', async () => {
		// A document that has been through a serializer is not guaranteed to come
		// back with the attributes written in the order they went out.
		const { html } = await renderDeck(SAMPLE_IR, NO_ASSETS)
		const reordered = html.replace(
			`<script type="application/json" id="${ISLAND_ID}">`,
			`<script id="${ISLAND_ID}" type="application/json">`
		)
		expect(reordered).not.toBe(html)
		expect(blocksOf(reordered).island).toBe(blocksOf(html).island)
	})
})
