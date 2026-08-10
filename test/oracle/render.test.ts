/**
 * The renderer against the real corpus.
 *
 * The claim under test is not "it produces HTML" — it is that **the island
 * survives whatever the visual channel does**. Every deck the importer can build
 * has to come back out of the document byte-identical, including the ones whose
 * picture this renderer is openly bad at, because that separation is what lets
 * the preview be approximate without the round trip being approximate.
 */

import { describe, expect, it } from 'vitest'
import { importDeck } from '../../src/import/deck'
import type { RenderIr } from '../../src/ir/render'
import { renderDeck } from '../../src/render/document'
import { ASSETS_ID, ISLAND_ID } from '../../src/render/island'
import { CORPUS, corpusBytes } from '../corpus/decks'

interface Imported {
	ir: RenderIr
	bytes: (name: string) => Uint8Array | undefined
}

const cache = new Map<string, Imported>()

async function importCorpus(name: string): Promise<Imported> {
	const hit = cache.get(name)
	if (hit) return hit
	const entry = CORPUS.find((candidate) => candidate.name === name)
	if (!entry) throw new Error(`the corpus has no deck named ${name}`)
	const { render, assets } = await importDeck(await corpusBytes(entry))
	const imported: Imported = { ir: render, bytes: (asset) => assets.bytesFor({ $asset: asset }) }
	cache.set(name, imported)
	return imported
}

function blockOf(html: string, id: string): string {
	const match = new RegExp(`<script type="application/json" id="${id}">(.*?)</script>`, 's').exec(html)
	if (match?.[1] === undefined) throw new Error(`no block with id ${id}`)
	return match[1]
}

describe('every corpus deck renders and comes back whole', () => {
	for (const entry of CORPUS) {
		it(`${entry.name}: the island parses back to the imported model`, async () => {
			const { ir, bytes } = await importCorpus(entry.name)
			const { html, warnings } = await renderDeck(ir, { bytes })

			expect(html.startsWith('<!doctype html>')).toBe(true)
			expect(JSON.parse(blockOf(html, ISLAND_ID))).toStrictEqual(ir)
			// One `<section>` per slide, whatever the slides contain.
			expect(html.match(/class="pxh-slide"/g)?.length ?? 0).toBe(ir.slides.length)
			// A warning is allowed (an unplaceable node, an oversized block); a
			// *missing asset* is not, because the importer supplies every one it
			// registered and a gap would mean the two disagree about the manifest.
			expect(warnings.filter((warning) => warning.includes('no bytes were supplied'))).toEqual([])
		})
	}
})

describe('the asset mode does not reach the model', () => {
	it('gives an inline and a ref document the same modelHash', async () => {
		// The property that keeps the oracle on one baseline instead of one per mode.
		// It holds because the island carries the manifest and never the bytes.
		const { ir, bytes } = await importCorpus('picture')
		const inline = await renderDeck(ir, { assets: 'inline', bytes })
		const ref = await renderDeck(ir, { assets: 'ref' })

		expect(inline.integrity).toStrictEqual(ref.integrity)
		expect(blockOf(inline.html, ISLAND_ID)).toBe(blockOf(ref.html, ISLAND_ID))
		expect(inline.html).toContain(`id="${ASSETS_ID}"`)
		expect(ref.html).not.toContain(`id="${ASSETS_ID}"`)
	})

	it('writes each asset into an inline document exactly once', async () => {
		// The tripwire for someone reintroducing `<img src="data:…">` at render
		// time: the bytes would then appear both in the asset block and on every
		// element that shows them, and a deck with one logo on ten slides would
		// carry eleven copies.
		const { ir, bytes } = await importCorpus('picture')
		const { html } = await renderDeck(ir, { assets: 'inline', bytes })

		expect(ir.assets.length).toBeGreaterThan(0)
		for (const asset of ir.assets) {
			const data = bytes(asset.name)
			if (data === undefined) throw new Error(`no bytes for ${asset.name}`)
			// A prefix rather than the whole payload: enough to be unique to these
			// bytes, short enough not to depend on base64 line handling.
			const prefix = btoa(String.fromCharCode(...data.subarray(0, 24)))
			expect(html.split(prefix).length - 1).toBe(1)
		}
		expect(html).not.toContain('src="data:')
	})
})

describe('what the picture admits to', () => {
	it('draws every geometry the corpus uses exactly, with no box standing in', async () => {
		// The corpus reaches `roundRect`, `rect`, `triangle`, `line` and a freeform,
		// and the local catalogue covers all five — so *zero* geometry fallbacks is
		// the real claim here, and it fails the day a deck is added that needs a
		// preset nobody has written the formula for. The marker's other half — that
		// an unresolved preset is visibly marked rather than passed off as exact —
		// is in `test/unit/render-geometry.test.ts`, where the preset can be chosen
		// rather than waited for.
		for (const entry of CORPUS) {
			const { ir, bytes } = await importCorpus(entry.name)
			const { html } = await renderDeck(ir, { bytes })
			expect([entry.name, html.match(/data-pxh-approx="geometry:[^"]+"/g) ?? []]).toStrictEqual([entry.name, []])
		}
	})

	it('draws an unstated fill as nothing and an unstated slide surface as white', async () => {
		// Two arms of `Fill.inherit` that have to disagree, and did not until they were
		// measured. A *shape* nothing states a fill for is unfilled, so painting one a
		// colour puts a rectangle behind text the deck wanted bare. A *slide* is the
		// surface everything sits on and always has one, so leaving it transparent
		// hands the deck whatever the embedding page happens to be.
		const { ir, bytes } = await importCorpus('master-background')
		const { html } = await renderDeck(ir, { bytes })

		expect(ir.slides[0]?.background.fill).toStrictEqual({ kind: 'inherit' })
		expect(html).toContain('fill="#ffffff" data-pxh-approx="background:inherit"')
		// The guess this replaced. Nothing may reintroduce a neutral for either arm.
		expect(html.toLowerCase()).not.toContain('d8dce6')
	})

	it('paints a baked shrink and leaves an unbaked one at full size', async () => {
		// The distinction the renderer used to collapse. `fontScale` and
		// `lnSpcReduction` are numbers the file states, so honouring them is reading
		// the deck; computing one for a frame that bakes none would be measuring, and
		// PowerPoint draws those at full size until the next edit anyway.
		const { ir, bytes } = await importCorpus('autofit-shrink')
		const { html } = await renderDeck(ir, { bytes })
		const frames = new Map(
			ir.slides[0]?.nodes.filter((node) => node.kind === 'shape').map((node) => [node.name, node]) ?? []
		)
		const styleOf = (name: string): string => {
			const id = frames.get(name)?.id
			const match = new RegExp(`data-pxh-run="${id}/0/0"[^>]*style="([^"]*)"`).exec(html)
			if (match?.[1] === undefined) throw new Error(`no first run for ${name}`)
			return match[1]
		}

		// All three frames are `shrink`; only two of them state what by.
		expect([...frames.keys()].map((name) => frames.get(name)?.text?.autofit)).toStrictEqual([
			'shrink',
			'shrink',
			'shrink',
		])
		expect(styleOf('baked')).toContain('font-size:19.6px') // 28pt × 70%
		expect(styleOf('spaced')).toContain('font-size:17.5px') // 28pt × 62.5%
		expect(styleOf('unbaked')).toContain('font-size:28px')

		// A stated percentage has the reduction subtracted from it, per ECMA-376
		// §21.1.2.1.3; an unstated one has no base, so it falls back and says so.
		expect(html).toContain('line-height:140%')
		expect(html).toContain('line-height:0.96')
		expect(html).toContain('data-pxh-approx="text:linespace"')
	})

	it('renders a chart as a labelled placeholder and never as a lookalike', async () => {
		const { ir, bytes } = await importCorpus('chart')
		const { html } = await renderDeck(ir, { bytes })
		expect(html).toContain('carried, not editable')
	})

	it('surfaces each slide’s declared differences in the page, not only the island', async () => {
		const { ir, bytes } = await importCorpus('table')
		const { html } = await renderDeck(ir, { bytes })
		const declared = ir.slides.flatMap((slide) => slide.fidelity)
		expect(declared.length).toBeGreaterThan(0)
		expect(html).toContain('Declared differences')
	})
})

describe('the editable surface reaches the document', () => {
	it('marks every run in surface as editable and addressable', async () => {
		// The parser reads these addresses back. If the renderer and `src/ir/surface.ts`
		// ever disagree about what is editable, the failure is silent in both
		// directions — an edit the page invited that the parser calls drift, or an
		// edit the parser expects that the page never offered.
		const { ir, bytes } = await importCorpus('text-box')
		const { html } = await renderDeck(ir, { bytes })

		const runs = ir.slides
			.flatMap((slide) => slide.nodes)
			.filter((node) => node.kind === 'shape')
			.flatMap((node) =>
				(node.text?.paragraphs ?? []).flatMap((paragraph, index) =>
					paragraph.runs.map((_, run) => `${node.id}/${index}/${run}`)
				)
			)

		expect(runs.length).toBeGreaterThan(0)
		for (const address of runs) expect(html).toContain(`data-pxh-run="${address}"`)
		expect(html.match(/contenteditable="true"/g)?.length ?? 0).toBe(runs.length)
	})
})
