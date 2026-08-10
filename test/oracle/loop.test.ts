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
import { IR_VERSION, type RenderIr } from '../../src/ir/render'
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

	it('carries underline and strike as the OOXML tokens, and reads them back', async () => {
		// The two properties the surface gained last, and the pair most able to fail
		// quietly: `RunProperties` spells them `none`/`single`/`double` and the deck
		// spells them `sng`/`dbl` and `sngStrike`/`dblStrike`, so an untranslated
		// value would be written verbatim, dropped by the writer as unrecognised, and
		// leave a document that looks edited and a deck that is not.
		//
		// Asserted twice over: on the emitted deck's own run options, and on a fresh
		// import of the emitted bytes — the second is what proves the loop closed
		// rather than that the right string was handed to the writer.
		const { source, html } = await rendered('text-box')
		const parsed = await parseDeck(html, { parseHtml: null })
		const shape = parsed.ir?.slides[0]?.nodes[0]
		if (shape?.kind !== 'shape' || !shape.text) throw new Error('fixture changed')
		const [first, second] = shape.text.paragraphs[0]?.runs ?? []
		if (!first || !second) throw new Error('the text-box deck lost a run')
		first.props.underline = 'single'
		second.props.strike = 'double'

		const emitted = await emitDeck(parsed, { source })
		const ir = readModelToIr(await Presentation.load(emitted.bytes))
		const runs = ir.slides[0]?.calls[0]?.args[0] as { options?: Record<string, unknown> }[]
		expect(runs[0]?.options?.underline).toStrictEqual({ style: 'sng' })
		expect(runs[1]?.options?.strike).toBe('dblStrike')

		const reimported = await importDeck(emitted.bytes)
		const back = reimported.render.slides[0]?.nodes[0]
		if (back?.kind !== 'shape') throw new Error('the emitted deck lost its text shape')
		expect(back.text?.paragraphs[0]?.runs[0]?.props.underline).toBe('single')
		expect(back.text?.paragraphs[0]?.runs[1]?.props.strike).toBe('double')
	})

	it('writes an explicit "not underlined", which is a different fact from saying nothing', async () => {
		// `none` is the value with no visible consequence and the most to lose: a run
		// that inherits `u="sng"` from its placeholder and states `u="none"` is not
		// underlined, and dropping the statement silently underlines it. So the
		// assertion is on the *presence* of the token, which a "nothing changed"
		// shortcut anywhere in the edit path would erase.
		const { source, html } = await rendered('text-box')
		const parsed = await parseDeck(html, { parseHtml: null })
		const shape = parsed.ir?.slides[0]?.nodes[0]
		if (shape?.kind !== 'shape' || !shape.text?.paragraphs[0]?.runs[0]) throw new Error('fixture changed')
		shape.text.paragraphs[0].runs[0].props.underline = 'none'
		shape.text.paragraphs[0].runs[0].props.strike = 'none'

		const emitted = await emitDeck(parsed, { source })
		const reimported = await importDeck(emitted.bytes)
		const back = reimported.render.slides[0]?.nodes[0]
		if (back?.kind !== 'shape') throw new Error('the emitted deck lost its text shape')
		const run = back.text?.paragraphs[0]?.runs[0]
		expect(run?.props.underline).toBe('none')
		expect(run?.props.strike).toBe('none')
	})

	it('carries a paragraph’s alignment, and puts it on every run of that paragraph', async () => {
		// The paragraph tier's first property, and the one place its mapping can fail
		// in a way no run-level test would catch. The writer groups the flat run list
		// into paragraphs and starts a new one wherever two adjacent runs disagree
		// about `align` — so writing the value onto the first run of a two-run
		// paragraph does not restyle it, it *splits* it. The run count is unchanged
		// either way, which is why the assertion is on the paragraph count.
		const { source, html } = await rendered('paragraph-align')
		const parsed = await parseDeck(html, { parseHtml: null })
		const shape = parsed.ir?.slides[0]?.nodes[0]
		if (shape?.kind !== 'shape' || !shape.text) throw new Error('the paragraph-align deck changed shape')
		const paragraphs = shape.text.paragraphs
		expect(paragraphs).toHaveLength(5)
		expect(paragraphs[4]?.runs).toHaveLength(2)

		// The last paragraph is the two-run one. Re-align it and the whole frame must
		// come back with the same five paragraphs.
		const target = paragraphs[4]
		if (!target) throw new Error('the paragraph-align deck lost its two-run paragraph')
		target.props.align = 'right'

		const emitted = await emitDeck(parsed, { source })
		const ir = readModelToIr(await Presentation.load(emitted.bytes))
		const runs = ir.slides[0]?.calls[0]?.args[0] as { options?: Record<string, unknown> }[]
		expect(runs.at(-2)?.options?.align).toBe('right')
		expect(runs.at(-1)?.options?.align).toBe('right')

		const reimported = await importDeck(emitted.bytes)
		const back = reimported.render.slides[0]?.nodes[0]
		if (back?.kind !== 'shape') throw new Error('the emitted deck lost its text shape')
		expect(back.text?.paragraphs).toHaveLength(5)
		expect(back.text?.paragraphs.map((paragraph) => paragraph.props.align)).toStrictEqual([
			'left',
			'center',
			'right',
			'justify',
			'right',
		])
	})

	it('returns a paragraph to inheriting its alignment, which states nothing rather than "left"', async () => {
		// The clearing half, and the one that separates *inherited* from *left*: the
		// writer omits `a:pPr/@algn` entirely when the option is absent, so a paragraph
		// whose alignment was cleared has to come back with no `align` at all. Writing
		// `left` instead would be a different paragraph — one that overrides whatever
		// its list style says rather than following it.
		const { source, html } = await rendered('paragraph-align')
		const parsed = await parseDeck(html, { parseHtml: null })
		const shape = parsed.ir?.slides[0]?.nodes[0]
		if (shape?.kind !== 'shape' || !shape.text?.paragraphs[1]) throw new Error('fixture changed')
		expect(shape.text.paragraphs[1].props.align).toBe('center')
		delete shape.text.paragraphs[1].props.align

		const emitted = await emitDeck(parsed, { source })
		const ir = readModelToIr(await Presentation.load(emitted.bytes))
		const runs = ir.slides[0]?.calls[0]?.args[0] as { options?: Record<string, unknown> }[]
		expect(runs[1]?.options && 'align' in runs[1].options).toBe(false)

		const reimported = await importDeck(emitted.bytes)
		const back = reimported.render.slides[0]?.nodes[0]
		if (back?.kind !== 'shape') throw new Error('the emitted deck lost its text shape')
		expect(back.text?.paragraphs[1]?.props.align).toBeUndefined()
	})

	it('clears a bullet to *inherited*, which is not the same deck as clearing it to none', async () => {
		// The load-bearing test of the paragraph tier's second property, and the one
		// ts-pptx#15 had to be fixed for. Absence is how this model spells inherited,
		// so clearing a bullet means deleting the key — but deleting the *option* means
		// `false`, which writes an explicit `<a:buNone/>` and overrides the master.
		// The two paint identically, so nothing but this assertion separates them:
		// `'inherit'` on the way out, and no bullet at all on the way back in.
		const { source, html } = await rendered('paragraph-bullet')
		const parsed = await parseDeck(html, { parseHtml: null })
		const shape = parsed.ir?.slides[0]?.nodes[0]
		if (shape?.kind !== 'shape' || !shape.text) throw new Error('the paragraph-bullet deck changed shape')
		const paragraphs = shape.text.paragraphs
		expect(paragraphs).toHaveLength(4)
		expect(paragraphs.map((paragraph) => paragraph.props.bullet?.kind)).toStrictEqual([
			undefined,
			'none',
			'character',
			undefined,
		])

		// Clear the glyph the third paragraph states. It must come back inheriting,
		// like the first, rather than explicitly bulletless, like the second.
		const target = paragraphs[2]
		if (!target) throw new Error('the paragraph-bullet deck lost its glyph paragraph')
		delete target.props.bullet

		const emitted = await emitDeck(parsed, { source })
		const ir = readModelToIr(await Presentation.load(emitted.bytes))
		const runs = ir.slides[0]?.calls[0]?.args[0] as { options?: Record<string, unknown> }[]
		expect(runs[2]?.options?.bullet).toBe('inherit')

		const reimported = await importDeck(emitted.bytes)
		const back = reimported.render.slides[0]?.nodes[0]
		if (back?.kind !== 'shape') throw new Error('the emitted deck lost its text shape')
		expect(back.text?.paragraphs).toHaveLength(4)
		expect(back.text?.paragraphs.map((paragraph) => paragraph.props.bullet?.kind)).toStrictEqual([
			undefined,
			'none',
			undefined,
			undefined,
		])
	})

	it('states a glyph on a two-run paragraph without splitting it in two', async () => {
		// The placement test, and the mirror image of the alignment one above it. A
		// paragraph property has to be written onto the runs `groupRunsIntoLines` will
		// read it from, and for `bullet` that is the *opening run alone*: a run that
		// states a glyph starts a new paragraph on its own, so the every-run placement
		// `align` needs would turn this two-run paragraph into two one-run paragraphs.
		// The run count does not move when it does, which is why the assertion is on
		// the paragraph count.
		//
		// It doubles as the code-point test. `Bullet` carries the character and the
		// option takes hex, so a value passed through untranslated would be written
		// verbatim, dropped by the writer as unrecognised, and leave a document that
		// looks edited beside a deck that is not.
		const { source, html } = await rendered('paragraph-bullet')
		const parsed = await parseDeck(html, { parseHtml: null })
		const shape = parsed.ir?.slides[0]?.nodes[0]
		if (shape?.kind !== 'shape' || !shape.text) throw new Error('the paragraph-bullet deck changed shape')
		const target = shape.text.paragraphs[3]
		if (!target) throw new Error('the paragraph-bullet deck lost its two-run paragraph')
		expect(target.runs).toHaveLength(2)
		target.props.bullet = { kind: 'character', char: '◆' }

		const emitted = await emitDeck(parsed, { source })
		const ir = readModelToIr(await Presentation.load(emitted.bytes))
		const runs = ir.slides[0]?.calls[0]?.args[0] as { options?: Record<string, unknown> }[]
		expect(runs.at(-2)?.options?.bullet).toStrictEqual({ characterCode: '25C6' })
		expect(runs.at(-1)?.options).not.toHaveProperty('bullet')

		const reimported = await importDeck(emitted.bytes)
		const back = reimported.render.slides[0]?.nodes[0]
		if (back?.kind !== 'shape') throw new Error('the emitted deck lost its text shape')
		expect(back.text?.paragraphs).toHaveLength(4)
		expect(back.text?.paragraphs[3]?.props.bullet).toMatchObject({ kind: 'character', char: '◆' })
		expect(back.text?.paragraphs[3]?.runs).toHaveLength(2)
	})

	it('suppresses an inherited bullet, which is the edit with something to show for it', async () => {
		// *Inherited* to *explicitly none*, the direction that changes the picture: the
		// first paragraph of this deck takes the master's glyph and has to stop, which
		// takes an `a:buNone` overriding the list style rather than the silence that
		// lets it through.
		const { source, html } = await rendered('paragraph-bullet')
		const parsed = await parseDeck(html, { parseHtml: null })
		const shape = parsed.ir?.slides[0]?.nodes[0]
		if (shape?.kind !== 'shape' || !shape.text?.paragraphs[0]) throw new Error('fixture changed')
		expect(shape.text.paragraphs[0].props.bullet).toBeUndefined()
		shape.text.paragraphs[0].props.bullet = { kind: 'none' }

		const emitted = await emitDeck(parsed, { source })
		const ir = readModelToIr(await Presentation.load(emitted.bytes))
		const runs = ir.slides[0]?.calls[0]?.args[0] as { options?: Record<string, unknown> }[]
		expect(runs[0]?.options?.bullet).toBe(false)

		const reimported = await importDeck(emitted.bytes)
		const back = reimported.render.slides[0]?.nodes[0]
		if (back?.kind !== 'shape') throw new Error('the emitted deck lost its text shape')
		expect(back.text?.paragraphs[0]?.props.bullet).toStrictEqual({ kind: 'none' })
	})

	it('carries a paragraph’s margins, and writes them on every run of it', async () => {
		// The paragraph tier's third and fourth properties. Nothing in the writer groups
		// on a margin, so the placement cannot split a paragraph the way `align`'s and
		// `bullet`'s can — the assertion is still on the paragraph count, because that
		// is what a *future* placement change would break first, and on both runs of the
		// two-run paragraph, which is the shape a fresh read produces.
		const { source, html } = await rendered('paragraph-indent')
		const parsed = await parseDeck(html, { parseHtml: null })
		const shape = parsed.ir?.slides[0]?.nodes[0]
		if (shape?.kind !== 'shape' || !shape.text) throw new Error('the paragraph-indent deck changed shape')
		const paragraphs = shape.text.paragraphs
		expect(paragraphs).toHaveLength(6)
		expect(paragraphs.map((paragraph) => paragraph.props.marginLeftPt)).toStrictEqual([undefined, 36, 36, 0, 36, 24])
		expect(paragraphs.map((paragraph) => paragraph.props.indentPt)).toStrictEqual([
			undefined,
			undefined,
			-18,
			undefined,
			-18,
			undefined,
		])

		const target = paragraphs[5]
		if (!target) throw new Error('the paragraph-indent deck lost its two-run paragraph')
		expect(target.runs).toHaveLength(2)
		target.props.marginLeftPt = 48
		target.props.indentPt = -12

		const emitted = await emitDeck(parsed, { source })
		const ir = readModelToIr(await Presentation.load(emitted.bytes))
		const runs = ir.slides[0]?.calls[0]?.args[0] as { options?: Record<string, unknown> }[]
		expect(runs.at(-2)?.options?.paraMarginLeft).toBe(48)
		expect(runs.at(-1)?.options?.paraMarginLeft).toBe(48)
		expect(runs.at(-2)?.options?.paraIndent).toBe(-12)

		const reimported = await importDeck(emitted.bytes)
		const back = reimported.render.slides[0]?.nodes[0]
		if (back?.kind !== 'shape') throw new Error('the emitted deck lost its text shape')
		expect(back.text?.paragraphs).toHaveLength(6)
		expect(back.text?.paragraphs[5]?.props).toMatchObject({ marginLeftPt: 48, indentPt: -12 })
		expect(back.text?.paragraphs[5]?.runs).toHaveLength(2)
	})

	it('returns a margin to inherited, which is not the same deck as setting it to zero', async () => {
		// The clearing half, and the one the whole property was waiting for. Deleting
		// the *option* does not state nothing — it restores the bullet's default, which
		// for anything but an inherited bullet is `marL="0"`, a margin the paragraph
		// never had. So the edit is made on the one paragraph whose bullet is explicitly
		// none: there, `'inherit'` and a deleted option produce different files, and the
		// stated zero above it is the control the wrong one would look identical to.
		const { source, html } = await rendered('paragraph-indent')
		const parsed = await parseDeck(html, { parseHtml: null })
		const shape = parsed.ir?.slides[0]?.nodes[0]
		if (shape?.kind !== 'shape' || !shape.text?.paragraphs[4]) throw new Error('fixture changed')
		expect(shape.text.paragraphs[4].props).toMatchObject({ marginLeftPt: 36, bullet: { kind: 'none' } })
		delete shape.text.paragraphs[4].props.marginLeftPt

		const emitted = await emitDeck(parsed, { source })
		const reimported = await importDeck(emitted.bytes)
		const back = reimported.render.slides[0]?.nodes[0]
		if (back?.kind !== 'shape') throw new Error('the emitted deck lost its text shape')
		expect(back.text?.paragraphs[4]?.props.marginLeftPt).toBeUndefined()
		// Everything else about that paragraph stands: the bullet it suppresses and the
		// indent it states are not what was cleared.
		expect(back.text?.paragraphs[4]?.props).toMatchObject({ indentPt: -18, bullet: { kind: 'none' } })
		// And the stated zero above it is untouched, and still stated.
		expect(back.text?.paragraphs[3]?.props.marginLeftPt).toBe(0)
	})

	it('keeps an inherited margin under a bullet it suppresses, which takes writing one edit as two', async () => {
		// The combination the two options unlocked, and the one place a paragraph edit
		// changes the meaning of an option it did not write: `bullet: false` writes
		// `marL="0" indent="0"` unless the margins say otherwise, and the contract omits
		// them for a paragraph whose bullet is inherited. So suppressing the bullet
		// alone would flatten the inherited indent in the same stroke, invisibly —
		// `pinMarginsBesideBullet` restates them as `'inherit'`, which is what a fresh
		// read of the result produces.
		const { source, html } = await rendered('paragraph-indent')
		const parsed = await parseDeck(html, { parseHtml: null })
		const shape = parsed.ir?.slides[0]?.nodes[0]
		if (shape?.kind !== 'shape' || !shape.text?.paragraphs[0]) throw new Error('fixture changed')
		const first = shape.text.paragraphs[0]
		expect(first.props.marginLeftPt).toBeUndefined()
		first.props.bullet = { kind: 'none' }

		const emitted = await emitDeck(parsed, { source })
		const reimported = await importDeck(emitted.bytes)
		const back = reimported.render.slides[0]?.nodes[0]
		if (back?.kind !== 'shape') throw new Error('the emitted deck lost its text shape')
		expect(back.text?.paragraphs[0]?.props.bullet).toStrictEqual({ kind: 'none' })
		expect(back.text?.paragraphs[0]?.props.marginLeftPt).toBeUndefined()
		expect(back.text?.paragraphs[0]?.props.indentPt).toBeUndefined()
	})

	it('refuses a margin the attribute cannot hold, rather than letting the writer clamp it', async () => {
		// The per-value bar on a measurement. `a:pPr/@marL` is unsigned, and the writer
		// clamps a negative one to zero and warns — which would move the text to a place
		// the caller did not ask for and the model does not hold. Refused here instead,
		// with the deck keeping the margin it had.
		const { source, html } = await rendered('paragraph-indent')
		const parsed = await parseDeck(html, { parseHtml: null })
		const shape = parsed.ir?.slides[0]?.nodes[0]
		if (shape?.kind !== 'shape' || !shape.text?.paragraphs[1]) throw new Error('fixture changed')
		shape.text.paragraphs[1].props.marginLeftPt = -10

		const emitted = await emitDeck(parsed, { source })
		expect(emitted.warnings.join('\n')).toContain('a margin of -10pt')

		const reimported = await importDeck(emitted.bytes)
		const back = reimported.render.slides[0]?.nodes[0]
		if (back?.kind !== 'shape') throw new Error('the emitted deck lost its text shape')
		expect(back.text?.paragraphs[1]?.props.marginLeftPt).toBe(36)
	})

	it('refuses a bullet the write API cannot author, and says so instead of rounding it', async () => {
		// The per-value bar, exercised from the one direction that can reach it: only
		// the delta is written, so a numbered bullet nobody touched is never
		// re-authored — but a caller editing the model directly can set one. The
		// contract is that the deck keeps what it had and the caller is told, which is
		// the same refusal the surface makes everywhere else, at value granularity.
		const { source, html } = await rendered('paragraph-bullet')
		const parsed = await parseDeck(html, { parseHtml: null })
		const shape = parsed.ir?.slides[0]?.nodes[0]
		if (shape?.kind !== 'shape' || !shape.text?.paragraphs[2]) throw new Error('fixture changed')
		shape.text.paragraphs[2].props.bullet = { kind: 'number', scheme: 'ea1ChsPeriod', startAt: 3 }

		const emitted = await emitDeck(parsed, { source })
		expect(emitted.warnings.join('\n')).toContain('"ea1ChsPeriod" numbered bullet')

		const reimported = await importDeck(emitted.bytes)
		const back = reimported.render.slides[0]?.nodes[0]
		if (back?.kind !== 'shape') throw new Error('the emitted deck lost its text shape')
		// The glyph the deck stated is still there — refused, not approximated.
		expect(back.text?.paragraphs[2]?.props.bullet).toMatchObject({ kind: 'character', char: '►' })
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
		// Off `IR_VERSION` rather than a literal: with the number written in, a bump
		// makes the replacement a no-op and the test still passes — asserting that an
		// unmodified document is refused, which it is not.
		const bumped = html.replace(`"irVersion":${IR_VERSION},"modelHash"`, '"irVersion":99,"modelHash"')
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
