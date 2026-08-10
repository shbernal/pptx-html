import { diffDeckIr, type FidelityNote } from '@shbernal/ts-pptx/script'
import { describe, expect, it } from 'vitest'
import { CORPUS, corpusBytes } from '../corpus/decks'
import { htmlLoop } from './html-lane'
import { assertRoundTrip, carryLoop, describeReport, resaveLoop, roundTrip, viewDeck } from './roundtrip'
import { scriptLoop } from './script-lane'

// Two things are under test here, and they are not the same thing:
//
//  1. the *lanes* — does a deck survive this path through the pipeline;
//  2. the *harness* — would it notice if a deck did not.
//
// (2) matters more. `diffDeckIr` is mutation-tested upstream, but the wrapper
// around it is ours and can silently compare the wrong two things: a harness
// that diffed a deck against itself, or that swallowed the notes argument, would
// report a clean run for every lane forever.

describe('lanes', () => {
	for (const entry of CORPUS) {
		it(`${entry.name}: survives a re-save`, async () => {
			const result = await roundTrip(await corpusBytes(entry), resaveLoop)
			assertRoundTrip(result, `${entry.name} (resave)`)
		})
	}

	for (const entry of CORPUS) {
		it(`${entry.name}: survives being carried across by import`, async () => {
			const result = await roundTrip(await corpusBytes(entry), carryLoop)
			assertRoundTrip(result, `${entry.name} (carry)`)
		})
	}

	for (const entry of CORPUS) {
		it(`${entry.name}: survives import → emit`, async () => {
			// Half the loop, with HTML nowhere near it. Both legs are upstream code,
			// so a failure here is an upstream issue to file rather than a renderer
			// bug — and every defect this catches is one that cannot later be
			// misattributed to the render or parse steps.
			const result = await roundTrip(await corpusBytes(entry), scriptLoop)
			assertRoundTrip(result, `${entry.name} (import → emit)`)
		})
	}

	for (const entry of CORPUS) {
		it(`${entry.name}: survives the whole loop through HTML`, async () => {
			// Invariant R itself. Everything above is a leg of this: the re-save lane
			// proves the harness, the carry lane proves the residual channel, the
			// import → emit lane proves the two upstream mappers — and this one adds
			// render and parse to the chain and asks the same question of all five.
			const result = await roundTrip(await corpusBytes(entry), htmlLoop)
			assertRoundTrip(result, `${entry.name} (html loop)`)
		})
	}

	it('carries a frame’s baked shrink through the loop, and keeps an unbaked one distinct', async () => {
		// This test used to assert the *loss*: `readModelToIr` reduced every
		// `normAutofit` to `fit: 'shrink'`, so a frame the writer emitted as
		// `<a:normAutofit fontScale="70000" lnSpcReduction="20000"/>` came back bare,
		// and `diffDeckIr` compared two `DeckIr`s both missing the field and called it
		// clean. Filed as https://github.com/shbernal/ts-pptx/issues/13 and pinned
		// inverted — asserting the loss rather than its absence — precisely so that
		// upstream fixing it would fail here rather than pass unnoticed. It did, so
		// the assertion is now the ordinary one.
		//
		// Kept as its own test rather than folded into the lanes above, because the
		// lanes could not see this and still could not: they assert that `diffDeckIr`
		// found nothing, which is only as strong as what the canonical model carries.
		// Reading the numbers off both views is what makes the coverage real.
		const entry = CORPUS.find((candidate) => candidate.name === 'autofit-shrink')
		if (!entry) throw new Error('corpus lost its autofit-shrink deck')
		const result = await roundTrip(await corpusBytes(entry), scriptLoop)

		const fitOf = (view: (typeof result)['input']): unknown[] =>
			view.ir.slides.flatMap((slide) => slide.calls.map((call) => (call.args[1] as { fit?: unknown })?.fit))
		// The object form and the string form are two states, not one value at two
		// precisions: ECMA-376 §21.1.2.1.3 defaults each attribute only when it is
		// *omitted*, and PowerPoint recomputes an unbaked scale on the next edit while
		// drawing a baked one exactly as written. Collapsing `unbaked` into
		// `{ fontScale: 100 }` would be the same bug in the other direction.
		const expected = [
			{ type: 'shrink', fontScale: 70, lnSpcReduction: 20 },
			{ type: 'shrink', fontScale: 62.5, lnSpcReduction: 10 },
			'shrink',
		]
		expect(fitOf(result.input)).toStrictEqual(expected)
		expect(fitOf(result.output)).toStrictEqual(expected)
		// And the canonical model carries them, which is the half that was silent:
		// without it the numbers could differ across the loop and the diff would agree.
		expect(JSON.stringify(result.input.canonical)).toContain('fontScale')
		expect(result.report.undeclared).toEqual([])
	})

	it('carries an explicit “not underlined” through the loop, distinct from saying nothing', async () => {
		// This test used to assert the *loss*, in the same inverted shape as the
		// autofit case above and for the same reason. `readModelToIr` mapped
		// `u="none"` and `strike="noStrike"` to `undefined`, so the token never
		// reached `DeckIr` and the re-emitted deck stated nothing — not the same fact,
		// since a run inheriting an underline from its list style and stating
		// `u="none"` is not underlined while the re-emitted one is. Filed as
		// https://github.com/shbernal/ts-pptx/issues/14 and pinned inverted so that a
		// fix would fail here rather than pass unnoticed. It did, so the assertion is
		// now the ordinary one.
		//
		// Still its own test rather than folded into the lanes, and the reason is
		// worth keeping: at the time it was written `canonicalDeckIr` omitted the
		// field too, so `diffDeckIr` compared two models *both* missing it and
		// `assertRoundTrip` passed on this deck. Reading the tokens off both views is
		// what made the coverage real then, and is what keeps it real now.
		const entry = CORPUS.find((candidate) => candidate.name === 'text-decoration')
		if (!entry) throw new Error('corpus lost its text-decoration deck')
		const result = await roundTrip(await corpusBytes(entry), scriptLoop)

		const decorationsOf = (view: (typeof result)['input']): unknown[] =>
			view.ir.slides.flatMap((slide) =>
				slide.calls.flatMap((call) =>
					(call.args[0] as { options?: { underline?: unknown; strike?: unknown } }[]).map(
						(run) => run.options?.underline ?? run.options?.strike ?? null
					)
				)
			)

		// All six states, in the deck's own order. The two `none`s are the ones that
		// were dropped; a run that genuinely states nothing would be `null` here, so
		// this array is also what distinguishes off from silent.
		const expected = [{ style: 'sng' }, { style: 'dbl' }, { style: 'none' }, 'sngStrike', 'dblStrike', 'noStrike']
		expect(decorationsOf(result.input)).toStrictEqual(expected)
		expect(decorationsOf(result.output)).toStrictEqual(expected)

		// And the half that was silent: the canonical model now carries the tokens, so
		// a future regression is a difference the lanes can see rather than one they
		// structurally cannot. Without this the values could drift and the diff agree.
		expect(JSON.stringify(result.input.canonical)).toContain('noStrike')
		// Nothing is lost on this path any more, so nothing declares a note for it
		// either — a note here now would be a stale declaration rather than a cover.
		expect(result.input.ir.fidelity.filter((note) => /underline|strike/i.test(note.construct))).toEqual([])
		expect(result.notes.filter((note) => /underline|strike/i.test(note.construct))).toEqual([])
		expect(result.report.undeclared).toEqual([])
	})

	it('the carry lane declares its gallery cost rather than hiding it', async () => {
		// The lane's note must actually be doing work: strip it and the same
		// run has to fail. A note that changes nothing is decoration.
		const result = await roundTrip(await corpusBytes(CORPUS[0]), carryLoop)
		expect(result.notes).toHaveLength(1)
		expect(result.report.declared.length).toBeGreaterThan(0)

		const undeclaredRun = diffDeckIr(result.input.canonical, result.output.canonical, [])
		expect(undeclaredRun.undeclared.length).toBeGreaterThan(0)
	})
})

describe('the harness itself', () => {
	async function canonicalFixture() {
		// A deck with a shape, text and explicit paint — enough surface to mutate.
		const entry = CORPUS.find((candidate) => candidate.name === 'autoshape')
		if (!entry) throw new Error('corpus lost its autoshape deck')
		return (await viewDeck(await corpusBytes(entry))).canonical
	}

	it('a deck compares clean against itself', async () => {
		const canonical = await canonicalFixture()
		expect(diffDeckIr(canonical, canonical, []).undeclared).toEqual([])
	})

	it('catches a changed option value', async () => {
		const expected = await canonicalFixture()
		const actual = structuredClone(expected)
		const call = actual.slides[0].calls[0]
		const options = call.args[call.args.length - 1] as Record<string, unknown>
		options.w = 99
		const report = diffDeckIr(expected, actual, [])
		expect(report.undeclared.length).toBeGreaterThan(0)
		expect(report.undeclared.some((difference) => difference.field === 'w')).toBe(true)
	})

	it('catches a dropped call', async () => {
		const expected = await canonicalFixture()
		const actual = structuredClone(expected)
		actual.slides[0].calls = []
		const report = diffDeckIr(expected, actual, [])
		expect(report.undeclared.length).toBeGreaterThan(0)
		expect(report.undeclared.some((difference) => difference.kind === 'lost')).toBe(true)
	})

	it('catches a slide rebound to the wrong layout', async () => {
		// The difference the projection exists to expose: a deck whose slides all
		// bind to the wrong layout emits identical *calls* and would otherwise
		// compare clean, while resolving theme and colour map against the wrong chrome.
		const expected = await canonicalFixture()
		const actual = structuredClone(expected)
		actual.slides[0].layoutName = 'NOT THE LAYOUT'
		expect(diffDeckIr(expected, actual, []).undeclared.length).toBeGreaterThan(0)
	})

	it('catches a lost slide', async () => {
		const expected = await canonicalFixture()
		const actual = structuredClone(expected)
		actual.slides = []
		expect(diffDeckIr(expected, actual, []).differences.length).toBeGreaterThan(0)
	})

	it('a matching note moves a difference from undeclared to declared', async () => {
		// The exclusion mechanism, end to end: the same mutation is a defect
		// without a note and the contract working with one. If this ever stops
		// holding, every "clean" run in the suite means nothing.
		const expected = await canonicalFixture()
		const actual = structuredClone(expected)
		const call = actual.slides[0].calls[0]
		const options = call.args[call.args.length - 1] as Record<string, unknown>
		// Only the width: a second changed field would be a second difference the
		// single note below does not cover, and the assertion would fail for a
		// reason that has nothing to do with note matching.
		options.line = { ...(options.line as Record<string, unknown>), width: 9 }

		const bare = diffDeckIr(expected, actual, [])
		const lineDifference = bare.undeclared.find((difference) => difference.field === 'width')
		expect(lineDifference, describeReport(bare)).toBeDefined()

		const note: FidelityNote = {
			slideNumber: lineDifference?.slideNumber ?? 1,
			shapeName: lineDifference?.shapeName ?? null,
			construct: 'line.width',
			disposition: 'dropped',
			cause: 'unwritable',
			detail: 'test note',
		}
		const declared = diffDeckIr(expected, actual, [note])
		expect(declared.undeclared).toHaveLength(0)
		expect(declared.declared.length).toBeGreaterThan(0)
	})

	it('a note for the wrong shape does not excuse the difference', async () => {
		// The failure mode a sloppier matcher would have: notes are scoped, so a
		// note about one shape must not silence another's defect.
		const expected = await canonicalFixture()
		const actual = structuredClone(expected)
		const call = actual.slides[0].calls[0]
		const options = call.args[call.args.length - 1] as Record<string, unknown>
		options.line = { ...(options.line as Record<string, unknown>), width: 9 }

		const note: FidelityNote = {
			slideNumber: 1,
			shapeName: 'a shape that is not on this slide',
			construct: 'line.width',
			disposition: 'dropped',
			cause: 'unwritable',
			detail: 'test note',
		}
		expect(diffDeckIr(expected, actual, [note]).undeclared.length).toBeGreaterThan(0)
	})

	it('two different decks do not compare clean', async () => {
		// The emptiness guard's cousin: proof the comparison is deck-sensitive at all.
		const first = (await viewDeck(await corpusBytes(CORPUS[0]))).canonical
		const second = (await viewDeck(await corpusBytes(CORPUS[3]))).canonical
		expect(diffDeckIr(first, second, []).undeclared.length).toBeGreaterThan(0)
	})
})
