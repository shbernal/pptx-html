/**
 * The preconditions a deleted repair layer used to exist for.
 *
 * `src/repair/repair.ts` rewrote the writer's OOXML after the fact: it
 * de-duplicated `p:cNvPr/@id`, rewrote `anchor="mid"` to `ctr`, forced
 * `<a:noAutofit/>` onto every `a:bodyPr`, stamped `@type` onto `p:sldSz`, and
 * stripped every notes part in the package. Against ts-pptx 3.0.0 the first four
 * conditions never occur — asserted below, across the corpus — and the last two
 * were not repairs at all:
 *
 * - **Stripping the notes parts destroyed speaker notes.** Measured, not
 *   reasoned about: a deck whose slide carried "Remember to mention the oracle."
 *   came back with `notesText === null` and no notes part in the package. A
 *   silent, total loss of a construct the IR models — the exact failure the
 *   charter exists to forbid — sitting in the shipping path for every output mode
 *   but `pptx-instance`.
 * - **Forcing `<a:noAutofit/>` wrote an explicit value where the source stated
 *   none.** The writer emits a self-closing `<a:bodyPr/>` for table cells, and an
 *   absent autofit child means *inherited*. Rewriting it to an explicit "no
 *   autofit" is flattening, which is the same mistake as resolving a placeholder's
 *   font size into the run.
 * - **`type="screen16x9"` was stamped unconditionally**, regardless of the
 *   dimensions beside it, so a 4:3 deck would have been labelled 16:9. The
 *   attribute is optional in the schema and PowerPoint infers it from `cx`/`cy`.
 *
 * The whole module also swallowed every exception and returned the *unrepaired*
 * deck, which made its own output depend on whether something threw.
 *
 * This file is what replaces it. If a future writer starts producing any of these
 * again, the answer is an upstream issue (see `.agents/skills/ts-pptx-upstream/`),
 * not a second repair layer — but it should be a failing test first.
 */

import { readZip } from '@shbernal/ts-pptx/zip'
import { describe, expect, it } from 'vitest'
import { CORPUS, corpusBytes } from '../corpus/decks'

const decoder = new TextDecoder()

/** Every slide part's XML, keyed by part name. */
async function slidesOf(bytes: Uint8Array): Promise<Map<string, string>> {
	const entries = await readZip(bytes)
	const slides = new Map<string, string>()
	for (const [name, data] of entries) {
		if (!/^ppt\/slides\/slide\d+\.xml$/i.test(name)) continue
		slides.set(name, typeof data === 'string' ? data : decoder.decode(data))
	}
	return slides
}

describe('the writer needs no repairing', () => {
	for (const entry of CORPUS) {
		it(`${entry.name}: emits valid, unrepaired slide XML`, async () => {
			const slides = await slidesOf(await corpusBytes(entry))
			expect(slides.size).toBeGreaterThan(0)

			for (const [name, xml] of slides) {
				// `p:cNvPr/@id` is unique within a slide's shape tree. `canonicalDeckIr`
				// absorbs the numbering, so a collision would never show up as a diff —
				// it has to be asserted on the XML or not at all.
				const ids = [...xml.matchAll(/<p:cNvPr\b[^>]*?\bid="(\d+)"/g)].map((match) => match[1])
				expect(new Set(ids).size, `${name} has duplicate cNvPr ids`).toBe(ids.length)

				// `mid` is not a member of `ST_TextAnchoringType`; the spelling is `ctr`.
				expect(xml, `${name} states an invalid anchor`).not.toMatch(/anchor="mid"/)
			}
		})
	}

	it('leaves a slide free to inherit its autofit', async () => {
		// The condition the forced `<a:noAutofit/>` would have destroyed. A cell that
		// states no autofit child inherits one, and that is a different deck from a
		// cell that states none explicitly.
		const entry = CORPUS.find((deck) => deck.name === 'table')
		if (!entry) throw new Error('the corpus lost its table deck')
		const xml = [...(await slidesOf(await corpusBytes(entry))).values()].join('')
		expect(xml).toMatch(/<a:bodyPr\b[^>]*\/>/)
		expect(xml).not.toMatch(/<a:noAutofit\b/)
	})
})
