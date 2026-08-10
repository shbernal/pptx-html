/**
 * pptx-html — public API.
 *
 * The package is a **loop**, and the four legs below are it:
 *
 * ```
 * .pptx  ──importDeck──►  RenderIr  ──renderDeck──►  HTML   (what a human edits)
 *   ▲                        ▲                        │
 *   └──emitDeck──────────────┴──────parseDeck─────────┘     (what a machine reads)
 * ```
 *
 * with one property to defend: for a deck written by `@shbernal/ts-pptx`,
 * `import → render → parse → emit` produces a deck equal to the input **under
 * the normalized read model**. Not byte-identical — zip order, `rId` numbering
 * and timestamps all vary legally. Slides whose features the model cannot
 * express are carried across intact rather than approximated, and nothing is
 * lost quietly: every element is **modeled**, **carried** or **warned**.
 *
 * The rendered HTML has two channels. The visible one (SVG and HTML) is allowed
 * to approximate — a preset geometry it cannot draw is marked and moved past.
 * The JSON island beside it is not, and it is the only channel `parseDeck`
 * trusts: the return path *parses the model*, it never re-derives it from
 * `getComputedStyle`. That is what makes an unresolved shape a cosmetic problem
 * rather than a lossy one.
 *
 * `emitDeck` needs the source package (`{ source }`), because masters, layouts,
 * theme and any carried slide's XML live there and nowhere else. The document
 * carries the edits; the caller supplies the substance.
 *
 * ## The heuristic lane
 *
 * `convertDeck` / `convertSlide` are the other lane: HTML that carries no
 * island, where the model is *inferred* from what the browser painted. It is
 * best-effort by construction and shares no code path with the loop above — see
 * `heuristic.SlideModel` for why its model is deliberately a different type. An
 * unmappable construct raises a `Warning` rather than being silently dropped.
 *
 * ## Where this runs
 *
 * `importDeck`, `renderDeck` and `emitDeck` are host-agnostic. `parseDeck` uses
 * a `ParentNode` when given one and falls back to scanning the document text
 * otherwise, so it works in Node with no DOM. The heuristic lane is
 * **browser-only**: it needs an iframe, `getComputedStyle`,
 * `getBoundingClientRect`, canvas and fonts.
 */

import {
	type ConvertOptions,
	convertDeck as convertDeckEngine,
	type ConvertResult,
	convertSlide as convertSlideEngine,
	type Warning,
} from './heuristic/engine'
import type { SlideModel } from './heuristic/model'

// ---------------------------------------------------------------------------
// The loop
// ---------------------------------------------------------------------------

export { type ImportedDeck, importDeck, importPresentation } from './import/deck'
export { renderDeck, type RenderedDeck, type RenderOptions } from './render/document'
export { type ParsedDeck, parseDeck, type ParseOptions } from './parse/deck'
export { emitDeck, type EmittedDeck, type EmitOptions } from './loop'

/**
 * The paint model the loop carries, and the handful of runtime values that go
 * with it: `IR_VERSION`, the EMU conversions and the node-id constructors. A
 * `export type *` here would emit those as *types* named after functions —
 * declared, unusable, and wrong in a way only a consumer would discover.
 */
export * from './ir/render'

/** The vocabulary of the return path: how a document is trusted, and which lane a slide took. */
export type { AssetResolver, ResolvedAssets } from './parse/assets'
export { IslandError, type IslandFault } from './parse/island'
export type { Lane, Reconciliation, SlideOutcome } from './parse/reconcile'
export type { Integrity } from './render/island'

/**
 * The editable surface: `project` is the part of a deck a rendered document may
 * change, `freeze` is its complement. Exported because "what may I safely edit
 * in this HTML?" is a question a caller has to be able to answer without reading
 * the renderer.
 */
export {
	EDITABLE_SURFACE,
	type EditableParaProp,
	type EditableRunProp,
	freeze,
	project,
	type SanctionedProjection,
} from './ir/surface'

// ---------------------------------------------------------------------------
// The heuristic lane
// ---------------------------------------------------------------------------

/**
 * The inference lane's DOM-shaped model, namespaced. It is *not* `RenderIr`, on
 * purpose (see `heuristic/model.ts`), and three of its type names — `Rect`,
 * `TableCell`, `Background` — mean something different from the loop's.
 * Namespacing it is what keeps the two vocabularies from being mistaken for one.
 */
export type * as heuristic from './heuristic/model'

export type { ConvertOptions, ConvertResult, OutputMode, ProgressEvent, Warning } from './heuristic/engine'
export type { IconResolver } from './heuristic/icons'

/**
 * Convert a full HTML document (head + slide sections) into PPTX.
 *
 * The heuristic lane's deck entry point: no island is read, the model is
 * inferred from the rendered page, and the result carries no round-trip
 * guarantee. Delivery is controlled by `opts.output` (default `'download'`);
 * `'base64'` / `'blob'` / `'pptx-instance'` return the deck on the result
 * instead. Icons resolve through `opts.resolveIcon`, and the writer through
 * `opts.pptxFactory`, both defaulted.
 */
export async function convertDeck(fullHtmlString: string, opts?: ConvertOptions): Promise<ConvertResult> {
	return convertDeckEngine(fullHtmlString, opts)
}

/**
 * Convert a single slide's HTML into the heuristic lane's slide model (plus
 * warnings). Per-slide entry point for testing and granularity: it renders the
 * slide and stops at the model, so that boundary can be asserted directly. Icons
 * resolve through `opts.resolveIcon`.
 */
export async function convertSlide(
	headHTML: string,
	slideHTML: string,
	opts?: ConvertOptions
): Promise<{ model: SlideModel; warnings: Warning[] }> {
	return convertSlideEngine(headHTML, slideHTML, opts)
}
