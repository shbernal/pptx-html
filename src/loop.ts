/**
 * The loop, composed: the four legs joined into one round trip.
 *
 * ```
 * .pptx ──importDeck──► IR ──renderDeck──► HTML ──parseDeck──► IR ──emitDeck──► .pptx
 * ```
 *
 * Each leg is documented where it lives; this file exists for the two decisions
 * that only make sense once they are joined.
 *
 * ## The document carries the edits; the package carries the deck
 *
 * `emitDeck` needs the source package, and that is a deliberate shape rather than
 * an unfinished one. Two things live there and nowhere else: the masters, layouts
 * and theme that `fromTemplate` preserves, and the verbatim XML of any carried
 * slide. Embedding a whole `.pptx` in the HTML to avoid asking for it would make
 * every document at least as large as the deck it shows, and it would make the
 * round trip depend on a copy of the package rather than on the package.
 *
 * This mirrors the decision one level down, where asset *bytes* are a render
 * option and never part of the model: the island states identity and the caller
 * supplies substance. Here the island states the edits and the caller supplies
 * the deck they apply to.
 *
 * ## The two must provably be the same deck
 *
 * An edit is addressed by node id, and an id means something only against the
 * shape tree it was assigned from. So before anything is applied, the package is
 * re-imported and its `modelHash` is compared with the document's. Equal means
 * the document was rendered from exactly these bytes — which also makes the check
 * a standing test that import is a function of the package and nothing else.
 * Unequal is a hard error: it is the one case where an edit could silently land
 * on a different shape.
 */

import { Presentation } from '@shbernal/ts-pptx/read'
import { emitDeckIr } from './emit/script'
import { importPresentation } from './import/deck'
import { applyEdits, editsBetween } from './parse/edits'
import { modelHashOf } from './parse/island'
import type { ParsedDeck } from './parse/deck'

export interface EmitOptions {
	/** The package the document was rendered from, as bytes. */
	source: Uint8Array
}

export interface EmittedDeck {
	bytes: Uint8Array
	/** Edits that could not be placed, and assets that could not be resolved. */
	warnings: string[]
}

/**
 * Emit the deck a parsed document describes.
 *
 * The source bytes are loaded twice on purpose, exactly as the oracle's lanes do:
 * `fromTemplate` strips the slides from the destination's copy, and `importSlide`
 * copies out of a live `Presentation`, so a carried slide needs a second one that
 * still has them.
 */
export async function emitDeck(parsed: ParsedDeck, options: EmitOptions): Promise<EmittedDeck> {
	if (parsed.ir === null || parsed.island === null || parsed.integrity === null) {
		throw new Error(
			'this document carries no IR island, so there is nothing to emit from; island-free HTML goes through the heuristic lane (convertDeck)'
		)
	}

	const source = await Presentation.load(options.source)
	const imported = await importPresentation(source)

	const hash = await modelHashOf(imported.render)
	if (hash !== parsed.integrity.modelHash) {
		throw new Error(
			`the document was not rendered from this package (modelHash ${parsed.integrity.modelHash.slice(0, 12)}… vs ${hash.slice(0, 12)}…); an edit addressed by node id cannot be placed against a different shape tree`
		)
	}

	const edits = editsBetween(parsed.island, parsed.ir)
	const applied = applyEdits(imported.deck, parsed.island, edits)

	return {
		bytes: await emitDeckIr(applied.deck, options.source, source),
		warnings: [...parsed.warnings, ...applied.warnings],
	}
}
