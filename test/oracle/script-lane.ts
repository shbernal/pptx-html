/**
 * The `import → emit` lane: half the loop, testable before any HTML exists.
 *
 * `readModelToIr` reads a deck into `DeckIr`; this replays that IR through the
 * write API and saves. Nothing renders, nothing parses, and both legs are
 * upstream code — which is the point. A difference here belongs to the
 * import/emit pair and cannot be misattributed to a renderer that has not been
 * written yet, and because neither leg is local, a failure is an upstream issue
 * to file rather than something to patch here.
 *
 * It is an **interpreter for the same tier `printScript` prints**, not a second
 * design. `printScript` emits `fromTemplate` + one generator per layout run +
 * `appendSlides`, and this makes the same calls directly, so `scriptTierNotes`
 * is the honest set of losses to declare (see `LoopOutput.notes`, where getting
 * this wrong both over- and under-excludes). Printing source and then evaluating
 * it would test a TypeScript emitter; this tests the mapping.
 *
 * The interpreter itself lives in `src/emit/script.ts`, as the package's emit
 * leg. What stays here is the *lane*: the deck it replays against
 * and the notes it declares, which are the oracle's business and not the
 * emitter's.
 */

import { emitDeckIr } from '../../src/emit/script'
import { type Loop, scriptTierNotes } from './roundtrip'

/**
 * The lane declares the tier's own losses and nothing else.
 *
 * It used to add one of its own — a table whose rows are **all** auto-height came
 * back with explicit heights, and `readModelToIr` excluded that case from its
 * `table.rowAuto` note explicitly. ts-pptx 3.0.0 fixed it
 * (https://github.com/shbernal/ts-pptx/issues/5), so `scriptTierNotes` now covers
 * it and a local compensation would be a second note for one loss.
 */
export const scriptLoop: Loop = async (input) => {
	const bytes = await emitDeckIr(input.ir, input.bytes, input.pres)
	return { bytes, notes: scriptTierNotes(input.ir) }
}
