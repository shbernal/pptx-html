import { Presentation } from '@shbernal/ts-pptx/read'
import { describe, expect, it } from 'vitest'
import { buildLabelledDeck, CORPUS, corpusBytes, labelledProducer } from '../corpus/decks'
import { first } from '../support'
import { htmlLoop } from './html-lane'
import { type DeckView, viewDeck } from './roundtrip'

// Two writer behaviours the rest of the plan leans on. Both are assumptions
// until something checks them, and both fail in ways that would otherwise be
// discovered late and blamed on the renderer.

/** Slide labels in deck order, read back out of the IR. */
function labels(view: DeckView): string[] {
	return view.ir.slides.map((slide) => {
		const text = slide.calls.find((call) => call.method === 'addText')
		const first = text?.args[0]
		return typeof first === 'string' ? first : (text?.sourceName ?? '?')
	})
}

describe('emit is a function of the deck, not of the run', () => {
	it('two writes of the same deck are equal under the normalized model', async () => {
		// Byte equality is explicitly not the claim: rIds, cNvPr ids, zip order and
		// timestamps all vary legally, and the canonicalizer absorbs exactly those.
		const entry = CORPUS[0]
		const first = await viewDeck(await entry.build())
		const second = await viewDeck(await entry.build())
		expect(second.canonical).toEqual(first.canonical)
	})

	it('the canonical model of a corpus deck matches its recorded form', async () => {
		// The cross-process half of the same claim. The snapshot was written by an
		// earlier process, so anything that varies per run — a module-level counter,
		// a clock, a map iteration order — fails here rather than passing a
		// same-process comparison that shares all of that state.
		const entry = CORPUS.find((candidate) => candidate.name === 'autoshape')
		if (!entry) throw new Error('corpus lost its autoshape deck')
		const view = await viewDeck(await corpusBytes(entry))
		expect(view.canonical).toMatchSnapshot()
	})

	it('a deck that has been through the whole loop matches its recorded form', async () => {
		// The same tripwire, on the four legs joined. `IR₂ ≡ IR₁ ⇒ pptx₂ ≡ pptx₁` is
		// what Invariant R rests on, and it is only true while emit is a function of
		// the IR — so the loop needs its own cross-process check rather than
		// inheriting the writer's. Anything the render or parse legs learn from the
		// host would show up here and nowhere else in the suite.
		const entry = CORPUS.find((candidate) => candidate.name === 'autoshape')
		if (!entry) throw new Error('corpus lost its autoshape deck')
		const input = await viewDeck(await corpusBytes(entry))
		const produced = await htmlLoop(input)
		expect((await viewDeck(produced.bytes)).canonical).toMatchSnapshot()
	})
})

describe('importSlide ordering', () => {
	it('appends at the current end, so import order is deck order', async () => {
		const source = await Presentation.load(await buildLabelledDeck(['one', 'two', 'three']))
		const destination = await Presentation.fromTemplate(await buildLabelledDeck([]))
		for (const index of [2, 0, 1]) destination.importSlide(source, index)
		expect(labels(await viewDeck(await destination.save()))).toEqual(['three', 'one', 'two'])
	})

	it('lets writer-authored and carried slides interleave in source order', async () => {
		// The shape the real pipeline produces: some slides re-emitted from the IR,
		// some carried across untouched, and the reader expecting them in the order
		// the source had them. There is no insertSlide/moveSlide, so ordering is a
		// consequence of call order — which is exactly why it is worth pinning.
		const source = await Presentation.load(await buildLabelledDeck(['carried-1', 'carried-2']))
		const destination = await Presentation.fromTemplate(await buildLabelledDeck([]))
		const layout = first(destination.layouts(), 'layout')

		destination.importSlide(source, 0)
		await destination.appendSlides(labelledProducer(['authored']), { layout })
		destination.importSlide(source, 1)

		expect(labels(await viewDeck(await destination.save()))).toEqual(['carried-1', 'authored', 'carried-2'])
	})
})
