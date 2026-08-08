/**
 * The scope threaded through every mapper: where we are in the source deck, what
 * we have found we cannot model, and how to reach the deck-wide tables a mapper
 * needs but must not own.
 *
 * A mapper that returned its notes would force every caller to collect and merge
 * them, and the merge is where a note quietly gets dropped. So notes accumulate
 * into one array the whole slide shares, and {@link forShape} narrows only the
 * *address* a note is stamped with.
 */

import type { Cause, Disposition, FidelityNote } from '@shbernal/ts-pptx/script'
import type { AssetIndex } from './assets'

export interface ImportScope {
	/** 1-based source slide. */
	slideNumber: number
	/** `p:cNvPr/@name` of the shape being mapped, or `null` above shape level. */
	shapeName: string | null
	/** Shared with every scope derived from this one. */
	notes: FidelityNote[]
	assets: AssetIndex
	/** Absolute partname of a slide → its 1-based number, for hyperlink targets. */
	slideNumberByPart: ReadonlyMap<string, number>
}

/** Narrow a scope to one shape, so its notes carry that shape's name. */
export function forShape(scope: ImportScope, shapeName: string): ImportScope {
	return { ...scope, shapeName }
}

/**
 * Record a loss.
 *
 * `construct` is an identifier matched mechanically, not a sentence — reuse an
 * upstream key (`text.align`, `line.dash`, `image.data`, …) whenever one names
 * the same construct, and coin a new dotted key only for something upstream does
 * not model at all. Coining a *synonym* is what makes the differ and the
 * renderer describe one loss two ways.
 */
export function note(
	scope: ImportScope,
	construct: string,
	disposition: Disposition,
	cause: Cause,
	detail: string
): void {
	scope.notes.push({
		slideNumber: scope.slideNumber,
		shapeName: scope.shapeName,
		construct,
		disposition,
		cause,
		detail,
	})
}
