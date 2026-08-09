/**
 * The loop's behaviour on inputs the corpus lanes cannot produce.
 *
 * `roundtrip.test.ts` asks one question of every deck — *did it survive?* — and
 * an unedited loop answers it without ever exercising the edit path, the
 * integrity checks or the lane decision. Those only appear when something is
 * deliberately wrong, or deliberately changed, which is what this file is for.
 *
 * The pairing matters: an assertion that no fallback happened is vacuous exactly
 * when the feature works, so each guarantee below is tested by *constructing* the
 * case rather than by observing its absence.
 */

import { Presentation } from '@shbernal/ts-pptx/read'
import { readModelToIr } from '@shbernal/ts-pptx/script'
import { describe, expect, it } from 'vitest'
import { importDeck, importPresentation } from '../../src/import/deck'
import type { RenderIr } from '../../src/ir/render'
import { emitDeck } from '../../src/loop'
import { parseDeck } from '../../src/parse/deck'
import { applyEdits, editsBetween } from '../../src/parse/edits'
import { IslandError } from '../../src/parse/island'
import { renderDeck } from '../../src/render/document'
import { ISLAND_ID } from '../../src/render/island'
import { CORPUS, corpusBytes } from '../corpus/decks'
import { viewDeck } from './roundtrip'

async function entry(name: string) {
	const found = CORPUS.find((candidate) => candidate.name === name)
	if (!found) throw new Error(`the corpus has no deck named ${name}`)
	return corpusBytes(found)
}

/** Render a corpus deck the way the loop does, bytes and all. */
async function rendered(name: string): Promise<{ source: Uint8Array; html: string }> {
	const source = await entry(name)
	const imported = await importDeck(source)
	const { html } = await renderDeck(imported.render, { bytes: (asset) => imported.assets.bytesFor({ $asset: asset }) })
	return { source, html }
}

/** Every run of text in an emitted deck, flattened, for asserting on content. */
async function textOf(bytes: Uint8Array): Promise<string[]> {
	const ir = readModelToIr(await Presentation.load(bytes))
	const texts: string[] = []
	const walk = (value: unknown): void => {
		if (Array.isArray(value)) {
			for (const item of value) walk(item)
			return
		}
		if (value === null || typeof value !== 'object') return
		const record = value as Record<string, unknown>
		if (typeof record.text === 'string') texts.push(record.text)
		for (const item of Object.values(record)) walk(item)
	}
	for (const slide of ir.slides) for (const call of slide.calls) walk(call.args)
	return texts
}

describe('an edit made through the surface', () => {
	it('reaches the emitted deck, and nothing beside it moves', async () => {
		// The claim the editable surface is for. Two runs share a text frame; one is
		// retyped, and the assertion is as much about the *other* one as about the
		// edited one — an edit placed by index against a misaligned call list would
		// overwrite its neighbour and still look like a success.
		const { source, html } = await rendered('text-box')
		const parsed = await parseDeck(html, { parseHtml: null })
		if (parsed.ir === null) throw new Error('the rendered document lost its island')

		const shape = parsed.ir.slides[0]?.nodes[0]
		if (shape?.kind !== 'shape' || !shape.text) throw new Error('the text-box deck changed shape')
		const target = shape.text.paragraphs[0]?.runs[0]
		if (!target) throw new Error('the text-box deck lost its first run')
		const original = target.text
		target.text = 'Edited'

		const emitted = await emitDeck(parsed, { source })
		const texts = await textOf(emitted.bytes)
		expect(texts).toContain('Edited')
		expect(texts).not.toContain(original)
		// The neighbouring run in the same frame is untouched.
		expect(texts).toContain('-trip')
	})

	it('carries a character property in the write API spelling', async () => {
		const { source, html } = await rendered('text-box')
		const parsed = await parseDeck(html, { parseHtml: null })
		const shape = parsed.ir?.slides[0]?.nodes[0]
		if (shape?.kind !== 'shape' || !shape.text?.paragraphs[0]?.runs[0]) throw new Error('fixture changed')
		shape.text.paragraphs[0].runs[0].props.sizePt = 40

		const emitted = await emitDeck(parsed, { source })
		const ir = readModelToIr(await Presentation.load(emitted.bytes))
		const runs = ir.slides[0]?.calls[0]?.args[0] as { options?: Record<string, unknown> }[]
		expect(runs[0]?.options?.fontSize).toBe(40)
	})

	it('drops a deleted node and leaves its siblings', async () => {
		// `bullet` has two shapes on one slide, so a deletion that took the wrong one
		// — or took both — is visible rather than indistinguishable from success.
		const { source, html } = await rendered('bullet')
		const parsed = await parseDeck(html, { parseHtml: null })
		if (parsed.ir === null) throw new Error('the rendered document lost its island')
		expect(parsed.ir.slides[0]?.nodes).toHaveLength(2)
		parsed.ir.slides[0]?.nodes.splice(0, 1)

		const emitted = await emitDeck(parsed, { source })
		const texts = await textOf(emitted.bytes)
		expect(texts).not.toContain('Deploy')
		expect(texts).toContain('Starred')
	})

	it('finds nothing to apply when nothing was edited', async () => {
		// The property the whole edit path is arranged around: an unedited document
		// exercises none of the structural mapping in `edits.ts`, so Invariant R for
		// the unedited loop does not depend on a single assumption in it. Asserted on
		// the delta itself — a table is the deepest structure the corpus has, so a
		// render or parse step that perturbed *anything* would show up here first.
		const { html } = await rendered('table')
		const parsed = await parseDeck(html, { parseHtml: null })
		if (parsed.island === null || parsed.ir === null) throw new Error('the rendered document lost its island')

		const edits = editsBetween(parsed.island, parsed.ir)
		expect(edits.runs.size).toBe(0)
		expect(edits.deleted.size).toBe(0)

		const applied = applyEdits(readModelToIr(await Presentation.load(await entry('table'))), parsed.island, edits)
		expect(applied.warnings).toEqual([])
	})
})

describe('what the return path refuses', () => {
	it('throws on a tampered island rather than falling back to inference', async () => {
		// The single most important refusal in the file. A model that has been
		// altered is the one input that must *not* be routed to the lane with the
		// weakest guarantees — which is what a fallback would do.
		const { html } = await rendered('autoshape')
		const tampered = html.replace('"rotation":0', '"rotation":45')
		expect(tampered).not.toBe(html)

		await expect(parseDeck(tampered, { parseHtml: null })).rejects.toThrow(IslandError)
		await expect(parseDeck(tampered, { parseHtml: null })).rejects.toThrow(/modelHash/)
	})

	it('throws when the document was rendered from a different package', async () => {
		// An edit is addressed by node id, and an id means nothing against another
		// deck's shape tree. Two different corpus decks both have `s1.sp2`.
		const { html } = await rendered('autoshape')
		const parsed = await parseDeck(html, { parseHtml: null })
		await expect(emitDeck(parsed, { source: await entry('gradient') })).rejects.toThrow(
			/not rendered from this package/
		)
	})

	it('refuses a version it cannot verify instead of guessing at the shape', async () => {
		const { html } = await rendered('text-box')
		const bumped = html.replace('"irVersion":2,"modelHash"', '"irVersion":99,"modelHash"')
		expect(bumped).not.toBe(html)
		await expect(parseDeck(bumped, { parseHtml: null })).rejects.toThrow(/IR version 99/)
	})

	it('refuses an island with no integrity block', async () => {
		const { html } = await rendered('text-box')
		const stripped = html.replace(/<script type="application\/json" id="pxh-integrity">.*?<\/script>/s, '')
		expect(stripped).not.toBe(html)
		await expect(parseDeck(stripped, { parseHtml: null })).rejects.toThrow(/no pxh-integrity block/)
	})

	it('rejects an asset whose bytes do not match the manifest', async () => {
		// The manifest's per-entry `sha256` earns its place here and nowhere else.
		// Flipping one byte is not a damaged picture, it is a different file, and
		// writing it into the deck under the original's name is a silent swap.
		const { html } = await rendered('picture')
		const imported = await importDeck(await entry('picture'))
		const name = imported.render.assets[0]?.name
		if (name === undefined) throw new Error('the picture deck carries no asset')

		const real = imported.assets.bytesFor({ $asset: name })
		if (real === undefined) throw new Error('the picture deck lost its bytes')
		const flipped = Uint8Array.from(real)
		flipped[flipped.length - 1] ^= 0xff

		await expect(parseDeck(html, { parseHtml: null, assets: () => flipped })).rejects.toThrow(/manifest hash/)
	})

	it('warns rather than throwing when an asset simply cannot be resolved', async () => {
		// The other side of the same rule. A `ref` document parsed without a resolver
		// has lost a picture, not its integrity, and failing the parse would trade the
		// thing this project guarantees for the thing it allows to be lossy.
		const imported = await importDeck(await entry('picture'))
		const { html } = await renderDeck(imported.render, { assets: 'ref' })
		const parsed = await parseDeck(html, { parseHtml: null })
		expect(parsed.ir).not.toBeNull()
		expect(parsed.assets.missing).toHaveLength(1)
		expect(parsed.warnings.join('\n')).toMatch(/could not be resolved/)
	})
})

describe('documents that are not ours', () => {
	it('reports the heuristic lane for island-free HTML instead of inventing a model', async () => {
		const parsed = await parseDeck('<!doctype html><html><body><h1>hand written</h1></body></html>', {
			parseHtml: null,
		})
		expect(parsed.ir).toBeNull()
		expect(parsed.slides).toEqual([])
		expect(parsed.warnings.join('\n')).toMatch(/heuristic lane/)
	})

	it('does not mistake the word for the block', async () => {
		// The string scanner's one real hazard: a document that mentions the id
		// without carrying the block. Reading it as an island would produce a parse
		// error on a page that simply is not one of ours.
		const parsed = await parseDeck(`<!doctype html><p>the island lives at id="${ISLAND_ID}"</p>`, { parseHtml: null })
		expect(parsed.ir).toBeNull()
	})
})

describe('determinism', () => {
	it('emits the same deck twice from the same document', async () => {
		// `emit(ir) ≡ emit(ir)` is what makes `IR₂ ≡ IR₁ ⇒ pptx₂ ≡ pptx₁`, and
		// therefore Invariant R, mean anything. Normalized equality, not bytes: `rId`
		// numbering and zip order vary legally and `canonicalDeckIr` absorbs them.
		const { source, html } = await rendered('gradient')
		const first = await emitDeck(await parseDeck(html, { parseHtml: null }), { source })
		const second = await emitDeck(await parseDeck(html, { parseHtml: null }), { source })

		const a = await viewDeck(first.bytes)
		const b = await viewDeck(second.bytes)
		expect(JSON.stringify(b.canonical)).toBe(JSON.stringify(a.canonical))
	})

	it('renders the same document from the same package twice', async () => {
		// Import feeding render has to be a function of the bytes alone, or the
		// `modelHash` check in `emitDeck` would reject documents at random.
		const source = await entry('nested-group')
		const once = await importPresentation(await Presentation.load(source))
		const twice = await importPresentation(await Presentation.load(source))
		expect(JSON.stringify(twice.render)).toBe(JSON.stringify(once.render))
	})

	it('produces the same island whether or not the bytes ride along', async () => {
		const imported = await importDeck(await entry('picture'))
		const inline = await renderDeck(imported.render, { bytes: (name) => imported.assets.bytesFor({ $asset: name }) })
		const ref = await renderDeck(imported.render, { assets: 'ref' })
		expect(ref.integrity).toStrictEqual(inline.integrity)

		const fromInline = await parseDeck(inline.html, { parseHtml: null })
		const fromRef = await parseDeck(ref.html, { parseHtml: null })
		expect(JSON.stringify(fromRef.island)).toBe(JSON.stringify(fromInline.island))
	})
})

describe('the island survives what a deck can contain', () => {
	it('round-trips a run whose text would end the script block', async () => {
		const source = await entry('text-box')
		const imported = await importDeck(source)
		const hostile: RenderIr = JSON.parse(JSON.stringify(imported.render))
		const shape = hostile.slides[0]?.nodes[0]
		if (shape?.kind !== 'shape' || !shape.text?.paragraphs[0]?.runs[0]) throw new Error('fixture changed')
		shape.text.paragraphs[0].runs[0].text = '</script><img src=x onerror=alert(1)>'

		const { html } = await renderDeck(hostile, { assets: 'ref' })
		const parsed = await parseDeck(html, { parseHtml: null })
		expect(parsed.island).toStrictEqual(hostile)
	})
})
