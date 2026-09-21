/**
 * The round-trip oracle: the machine that answers *did the loop preserve this
 * deck?* for any deck and any lane through the pipeline.
 *
 * It defines no equality of its own. Normalized-model equality is already
 * implemented upstream, and this file is the harness around it:
 *
 * ```
 * .pptx bytes ──Presentation.load──► read model ──readModelToIr──► DeckIr
 *                                                     │
 *                                            canonicalDeckIr
 *                                                     ▼
 *                             diffDeckIr(expected, actual, notes) ──► RoundTripReport
 * ```
 *
 * Byte equality is the wrong target: zip entry order, `dcterms:created`, `rId`
 * numbering and `cNvPr` ids all differ legally between two correct writes of the
 * same deck, and `canonicalDeckIr` absorbs exactly those. `RoundTripReport`
 * already carries the severity split — `undeclared` is the defect count and the
 * only field a gate may look at; `declared` and `added` are the contract working.
 *
 * A *lane* is any path from bytes back to bytes (see {@link Loop}). Today two
 * exist — re-save and carry-by-import — and the HTML render/parse lane plugs in
 * behind the same seam without changing a line here.
 */

import { Presentation } from '@shbernal/ts-pptx/read'
import {
	type CanonicalDeck,
	canonicalDeckIr,
	type DeckIr,
	diffDeckIr,
	type FidelityNote,
	type IrDifference,
	printScript,
	type RoundTripReport,
	readModelToIr,
} from '@shbernal/ts-pptx/script'

/** One deck, at every level the oracle needs to see it. */
export interface DeckView {
	bytes: Uint8Array
	pres: Presentation
	ir: DeckIr
	canonical: CanonicalDeck
}

/** Load a deck and reduce it to the comparable form. */
export async function viewDeck(bytes: Uint8Array): Promise<DeckView> {
	const pres = await Presentation.load(bytes)
	const ir = readModelToIr(pres)
	return { bytes, pres, ir, canonical: canonicalDeckIr(ir) }
}

/**
 * What a lane returns: the deck it produced, plus **the losses that lane
 * declares**.
 *
 * `notes` is the load-bearing field and the easiest thing in the whole harness
 * to get wrong. `diffDeckIr` excludes exactly the differences its `notes`
 * account for, so notes belonging to a *different* output tier both over- and
 * under-exclude: they excuse defects this lane really has and leave differences
 * it legitimately declared reported as bugs. Upstream states the rule for its
 * own printer (use `printScript`'s returned notes, never `DeckIr.fidelity`); the
 * general form is **the notes of the tier that produced these bytes**, which is
 * why they are returned by the lane rather than derived here.
 *
 * Omitting `notes` therefore means something specific: *this lane claims to lose
 * nothing*. Both lanes below claim exactly that.
 */
export interface LoopOutput {
	bytes: Uint8Array
	notes?: FidelityNote[]
}

/** A path from a deck back to a deck. The unit the oracle judges. */
export type Loop = (input: DeckView) => Promise<LoopOutput>

export interface RoundTripResult {
	input: DeckView
	output: DeckView
	/** The notes the lane declared; `[]` for a lane claiming losslessness. */
	notes: FidelityNote[]
	report: RoundTripReport
}

/** Run one deck through one lane and diff what came back against what went in. */
export async function roundTrip(bytes: Uint8Array, loop: Loop): Promise<RoundTripResult> {
	const input = await viewDeck(bytes)
	const produced = await loop(input)
	const output = await viewDeck(produced.bytes)
	const notes = produced.notes ?? []
	return {
		input,
		output,
		notes,
		report: diffDeckIr(input.canonical, output.canonical, notes),
	}
}

/**
 * The notes that apply to upstream's own template-anchored script tier.
 *
 * Here so that the one case where `printScript`'s notes *are* the right answer
 * has a name, and so nothing reaches for `DeckIr.fidelity` by accident. Do not
 * pass this from a lane that is not the printed script — see {@link LoopOutput}.
 */
export function scriptTierNotes(ir: DeckIr): FidelityNote[] {
	return printScript(ir, { assets: 'inline' }).notes
}

/** How much deck there actually is, for the emptiness guard below. */
export interface DeckWeight {
	slides: number
	calls: number
	authored: number
	carried: number
}

export function weigh(view: DeckView): DeckWeight {
	const slides = view.ir.slides
	return {
		slides: slides.length,
		calls: slides.reduce((n, slide) => n + slide.calls.length, 0),
		authored: slides.filter((slide) => slide.source === 'authored').length,
		carried: slides.filter((slide) => slide.source === 'carried').length,
	}
}

/**
 * Guard against the oracle's most embarrassing failure mode: an import that
 * yields empty slides compares clean against another empty import and passes
 * every test in the suite. A deck the oracle is asked about must have content.
 */
export function assertNonTrivial(view: DeckView, label: string): void {
	const weight = weigh(view)
	if (weight.slides === 0) throw new Error(`${label}: deck has no slides — an empty deck proves nothing`)
	if (weight.calls === 0)
		throw new Error(`${label}: deck has ${weight.slides} slide(s) but no calls — nothing was read`)
}

function describeDifference(difference: IrDifference): string {
	const shape = difference.shapeName ? ` (${difference.shapeName})` : ''
	return `  slide ${difference.slideNumber}${shape} ${difference.path} [${difference.kind}] expected ${difference.expected}, got ${difference.actual}`
}

/** Human-readable summary of a report, for a failure message or the coverage log. */
export function describeReport(report: RoundTripReport): string {
	const lines = [
		`${report.slideCount} slide(s): ${report.undeclared.length} undeclared, ${report.declared.length} declared, ${report.added.length} added, ${report.unmatchedNotes.length} unmatched note(s)`,
	]
	for (const difference of report.undeclared) lines.push(describeDifference(difference))
	return lines.join('\n')
}

/**
 * The gate. `undeclared` is the only field that fails a run:
 *
 * - `declared` — a difference a note predicted; the fidelity contract working.
 * - `added` — the write path spelling out what the source left implicit.
 * - `unmatchedNotes` — usually a construct the read model cannot see at all, so
 *   it is absent from *both* sides and has nothing to match. Reported, never
 *   gated on, and never "fixed" by deleting the note.
 */
export function assertRoundTrip(result: RoundTripResult, label: string): void {
	if (result.report.undeclared.length === 0) return
	throw new Error(`${label}: round-trip lost something no fidelity note declared.\n${describeReport(result.report)}`)
}

/**
 * Lane: re-save the package.
 *
 * The weakest possible loop, and it is here to test the *harness* rather than
 * the pipeline: it exercises load → IR → canonicalize on two genuinely different
 * byte streams, so a harness that compared a deck with itself by accident, or
 * whose canonicalization was position-dependent, fails here rather than passing
 * everything.
 */
export const resaveLoop: Loop = async (input) => ({ bytes: await input.pres.save() })

/**
 * Lane: carry every slide across by import — the residual channel of Invariant R.
 *
 * `fromTemplate` keeps the source's masters, layouts and theme byte-identical and
 * strips its slides; `importSlide` then brings each slide back with its own XML
 * intact. This is the lane an unmodeled slide takes through the real pipeline, so
 * it is worth gating on now, before anything renders.
 *
 * Note this needs the source *package*, which is why it is `importSlide` and not
 * `appendSlides`: the latter takes a serializable `SlideSource` and can be fed
 * from bytes embedded in HTML, but it costs placeholder inheritance and
 * re-resolves `schemeClr` against the destination theme.
 *
 * **This lane used to declare a cost and no longer does.** `importSlide`'s
 * `theme: 'copy'` default brought the slide's own layout → master → theme
 * subgraph across unconditionally, so the destination ended up holding the
 * source's master beside its own and the layout gallery gained an entry nothing
 * bound to — declared here as a `master.default` note. As of `@shbernal/ts-pptx`
 * 3.5.0 the import reuses chrome the destination already holds, and this lane
 * templates the destination *from the source*, so there is nothing left to
 * duplicate. Verified rather than assumed: across the whole corpus the note
 * matched no difference on any deck, and dropping it moves `undeclared` on none
 * of them.
 *
 * The note is gone rather than kept, because a note that excludes nothing is the
 * decoration the lane test exists to catch — and a stale one is worse than none,
 * since it claims a loss the loop no longer takes. The test below now pins the
 * clean run, so a regression that reintroduces the duplicate arrives as an
 * `undeclared` difference instead of being absorbed by a note left behind to
 * cover it.
 */
export const carryLoop: Loop = async (input) => {
	const destination = await Presentation.fromTemplate(input.bytes)
	// One batch, not a loop of `importSlide`. A slide that links to another slide in the
	// same deck can only be copied alongside its target: ts-pptx 4.0 refuses a source page
	// whose link leaves the set being imported, because the alternative is rewriting the
	// relationship to point at a page that is not there. Imported one at a time, every page
	// is its own set, so the `hyperlink` deck's slide-to-slide link had nowhere to land; the
	// batch makes the whole deck the set and the link resolves inside it.
	destination.importSlides(
		input.pres.slides.map((_, index) => ({
			source: input.pres,
			sourceIndex: index,
			outputIndex: index,
			importNotes: true,
		}))
	)
	return { bytes: await destination.save() }
}
