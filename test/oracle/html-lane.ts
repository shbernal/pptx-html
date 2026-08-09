/**
 * The HTML lane: the whole loop, behind the oracle's existing seam.
 *
 * ```
 * bytes ──import──► RenderIr ──render──► HTML ──parse──► RenderIr ──emit──► bytes
 * ```
 *
 * This is the lane Invariant R is actually about, and it plugs in as a `Loop`
 * exactly like the two that came before it — which was the point of judging lanes
 * rather than one fixed pipeline.
 *
 * ## It declares the script tier's notes, and that is not a shortcut
 *
 * The final leg is `emitDeckIr`, the same interpreter the `import → emit` lane
 * uses, so the losses this lane can have are that tier's losses. The render and
 * parse legs in between are *lossless by construction*: the island is the model
 * verbatim, and it comes back byte-identical or the parse throws. So adding notes
 * of its own here would excuse defects the lane really has — `LoopOutput.notes`
 * spells out why that is the easiest thing in the harness to get wrong.
 *
 * If that claim is ever false, this lane fails rather than passing quietly: the
 * island round-trip is asserted separately in `render.test.ts`, and `parseDeck`
 * throws on any mismatch rather than degrading.
 *
 * ## No DOM, on purpose
 *
 * The oracle runs in Node, so `parseDeck` is given the rendered string with no
 * HTML parser. That reads the island and skips the visual channel, which is the
 * correct behaviour for a document that was never opened, let alone typed into —
 * and it keeps this lane measuring the *loop* rather than a browser's
 * contenteditable. Edits are exercised where they live: `test/unit/parse-*` for
 * the decision logic, `test/browser/` for the reading.
 */

import { importPresentation } from '../../src/import/deck'
import { emitDeck } from '../../src/loop'
import { parseDeck } from '../../src/parse/deck'
import { renderDeck } from '../../src/render/document'
import type { Loop } from './roundtrip'
import { scriptTierNotes } from './roundtrip'

/** Render with the bytes inlined, then read the document back and emit it. */
export const htmlLoop: Loop = async (input) => {
	const imported = await importPresentation(input.pres)
	const { html } = await renderDeck(imported.render, { bytes: (name) => imported.assets.bytesFor({ $asset: name }) })

	// `parseHtml: null` rather than relying on the absence of `DOMParser`: the
	// oracle's environment is a fact about the test runner, and a lane whose
	// behaviour changes when that changes is not measuring what it claims to.
	const parsed = await parseDeck(html, { parseHtml: null })
	const emitted = await emitDeck(parsed, { source: input.bytes })

	return { bytes: emitted.bytes, notes: scriptTierNotes(input.ir) }
}
