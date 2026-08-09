/**
 * HTML → {@link RenderIr}: the third leg of the loop.
 *
 * ```
 * .pptx ──import──► RenderIr ──render──► HTML ──this file──► RenderIr ──emit──► .pptx
 * ```
 *
 * The whole of the return path's judgement is in three lines of this file: read
 * the island, read the surface, fold one into the other. Everything else is
 * delegated, and each piece is documented where it lives — `island.ts` for what
 * makes a model believable, `assets.ts` for what makes bytes believable,
 * `surface.ts` for what the DOM is allowed to say, `reconcile.ts` for what an
 * edit is allowed to mean.
 *
 * ## Why the surface needs a DOM and the island does not
 *
 * The island is inert text and a string scanner reads it exactly, in Node or in a
 * browser. The *edits* are not text — they are the state of a live contenteditable
 * DOM — so reading them needs a real document. A caller in a browser passes one
 * (`document`, or a `DOMParser` result) and gets the full four-lane behaviour. A
 * caller in Node passes a string; the island is read, the document's visual
 * channel is not, and that is stated rather than assumed away. It is the honest
 * reading for the case it is actually for: a string that has never been a live
 * document cannot have been typed into.
 */

import type { RenderIr } from '../ir/render'
import type { Integrity } from '../render/island'
import { type AssetResolver, type ResolvedAssets, resolveAssets } from './assets'
import { blocksOf, openIsland } from './island'
import { type Reconciliation, reconcile, type SlideOutcome } from './reconcile'
import { readSurface } from './surface'

export interface ParseOptions {
	/**
	 * Where asset bytes come from. Consulted before the document's own inline
	 * block, and every byte it returns is checked against the manifest hash.
	 */
	assets?: AssetResolver
	/**
	 * Parse a string source into a DOM so the surface can be read from it.
	 * Defaults to the platform's `DOMParser` where there is one. Supply it
	 * explicitly to use a parser in Node, or pass `null` to read the island only.
	 */
	parseHtml?: ((html: string) => ParentNode) | null
}

export interface ParsedDeck {
	/**
	 * The island's model with the document's sanctioned edits folded in, or `null`
	 * when the document carried no island at all.
	 */
	ir: RenderIr | null
	/** The model exactly as it was rendered, before any edit. `modelHash` covers this. */
	island: RenderIr | null
	integrity: Integrity | null
	/** One entry per slide, naming the lane it took. Empty on the heuristic lane. */
	slides: SlideOutcome[]
	assets: ResolvedAssets
	warnings: string[]
}

function defaultParser(): ((html: string) => ParentNode) | null {
	if (typeof DOMParser === 'undefined') return null
	return (html) => new DOMParser().parseFromString(html, 'text/html')
}

/**
 * Read a rendered document back into the model it was rendered from.
 *
 * Throws — never falls back — when the document has an island it cannot trust:
 * a hash mismatch, a version this build does not read, a block that is not JSON.
 * See `island.ts` for why a tampered model must not be routed to the lane with
 * the weakest guarantees.
 */
export async function parseDeck(source: string | ParentNode, options: ParseOptions = {}): Promise<ParsedDeck> {
	const blocks = blocksOf(source)
	const opened = await openIsland(blocks)

	// No island: this document was not rendered by `renderDeck`, so there is no
	// model to read and nothing here will invent one. The heuristic lane is a
	// *different entry point* (`convertDeck`), not a fallback inside this one —
	// see the trap it exists to avoid in `heuristic/model.ts`. Returning `ir: null`
	// rather than an inferred model is what keeps a caller from receiving a
	// guessed deck through the function whose whole contract is that it does not
	// guess.
	if (opened === null) {
		return {
			ir: null,
			island: null,
			integrity: null,
			slides: [],
			assets: { bytesFor: () => undefined, missing: [], warnings: [] },
			warnings: [
				'the document carries no IR island, so there is no model to read back; this is the heuristic lane, which is `convertDeck` and makes no round-trip guarantee',
			],
		}
	}

	const assets = await resolveAssets(opened.ir.assets, { resolver: options.assets, inline: blocks.assets })
	const warnings = [...assets.warnings]

	const root = rootOf(source, options, warnings)
	const applied: Reconciliation =
		root === null
			? // A *copy*, never the island itself. `island` is what `modelHash` covers
				// and what an edit is diffed against, so handing back one object under
				// two names would make every edit invisible to `editsBetween` — the
				// baseline would move with the change.
				{ ir: structuredClone(opened.ir), slides: opened.ir.slides.map(untouched), warnings: [] }
			: reconcile(opened.ir, readSurface(root, opened.ir))
	warnings.push(...applied.warnings)

	return { ir: applied.ir, island: opened.ir, integrity: opened.integrity, slides: applied.slides, assets, warnings }
}

function untouched(slide: RenderIr['slides'][number]): SlideOutcome {
	return { number: slide.number, lane: 'exact', edits: 0, notes: [] }
}

/**
 * The DOM to read edits from, or `null` when there is none to read.
 *
 * The warning in the last branch is the important part. Reporting every slide
 * `exact` is *true* — the deck that comes out equals the island that went in —
 * but it is true for a reason a caller has to know, because any edit made to the
 * document has just been dropped on the floor.
 */
function rootOf(source: string | ParentNode, options: ParseOptions, warnings: string[]): ParentNode | null {
	if (typeof source !== 'string') return source
	const parse = options.parseHtml === undefined ? defaultParser() : options.parseHtml
	if (parse !== null) return parse(source)
	warnings.push(
		'no HTML parser is available, so the document was read from its island alone; every slide is reported exact because the emitted deck matches the island, and any edit made to the document has been discarded'
	)
	return null
}
