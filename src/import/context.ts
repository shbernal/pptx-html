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

import type { Cause, Disposition, FidelityNote } from 'pptx-ts/script'
import type { ChromeTier } from '../ir/render'
import type { AssetIndex } from './assets'

export interface ImportScope {
	/** 1-based source slide. */
	slideNumber: number
	/** `p:cNvPr/@name` of the shape being mapped, or `null` above shape level. */
	shapeName: string | null
	/**
	 * Which shape tree is being walked: `null` for the slide's own, otherwise the
	 * template tier the shapes were inherited from. It selects the id namespace
	 * (see `chromeNodeId`) and nothing else — every mapper below is the same code
	 * for all three trees, which is the point of upstream returning one `AnyShape`
	 * union from all of them.
	 */
	chrome: ChromeTier | null
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
 * Re-point a scope at a template tier's shape tree.
 *
 * The notes array is **replaced, not shared**, and that is the whole reason this
 * is a function rather than a spread at the call site. `FidelityNote` is the
 * round-trip vocabulary: it says what the emitted deck lost. Chrome is not
 * emitted — the layout and master parts ride along in the template package — so a
 * note filed while walking them would report a loss the round trip never takes,
 * on a slide whose own content is intact. The notes are still *collected* so a
 * mapper need not know where it is; they are simply not the slide's.
 */
export function forChrome(scope: ImportScope, tier: ChromeTier): ImportScope {
	return { ...scope, chrome: tier, shapeName: null, notes: [] }
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
