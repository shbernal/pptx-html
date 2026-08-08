/**
 * dom2pptx — public API.
 *
 * Turns HTML from the DOM into editable PPTX: it builds an intermediate slide
 * model (see `./ir/model`) from the rendered DOM and emits it through
 * `@shbernal/ts-pptx`. This package owns the HTML → ts-pptx link only; it does
 * not write OOXML itself, and it does not generate the HTML.
 *
 * Every element is **modeled**, **carried** or **warned** — never approximated
 * into something that cannot be read back. This lane (HTML that carries no IR)
 * is the inference lane, so its mapping is heuristic: an unmappable construct
 * raises a `Warning` rather than being silently dropped or rasterized.
 *
 * This is a **browser** package: it needs a real DOM (iframe, `getComputedStyle`,
 * `getBoundingClientRect`, canvas, fonts) and is not Node-portable as written.
 *
 * `convertDeck` runs the full engine (`./engine`, split across `extract` /
 * `emit` / `repair`); `convertSlide` is the per-slide entry that stops at the IR
 * so the model boundary can be asserted directly.
 */

import { convertDeck as convertDeckEngine, convertSlide as convertSlideEngine } from './engine'
import type { SlideModel } from './ir/model'

// The legacy IR. `./ir/render` (`RenderIr`) and `./import` are deliberately
// *not* re-exported here yet: two of the model's type names (`TableCell`,
// `Background`) collide with the legacy ones below, and an `importDeck` whose
// return type cannot be named is worse than none. Both problems disappear in the
// same edit — when this line goes and `RenderIr` takes its place.
export type * from './ir/model'

/** How a non-fatal issue surfaced during conversion. */
export interface Warning {
	slide?: number
	code?: string
	message: string
}

export interface ProgressEvent {
	phase: 'parse' | 'render' | 'extract' | 'emit' | 'finalize'
	completed?: number
	total?: number
	stopped?: boolean
}

/** Resolve an iconify icon name (e.g. `mdi:home`) to an SVG string, or null. */
export type IconResolver = (name: string) => Promise<string | null>

/** Where the converted deck should be delivered. */
export type OutputMode = 'download' | 'base64' | 'blob' | 'pptx-instance'

export interface ConvertOptions {
	author?: string
	title?: string
	lang?: string
	/** Used for `output: 'download'`. */
	fileName?: string
	fontConfig?: Record<string, unknown>
	onProgress?: (event: ProgressEvent) => void
	/** Cooperative cancellation: return true to stop after the current slide. */
	shouldStop?: () => boolean
	/**
	 * Icon resolver seam. Defaults to fetching from `api.iconify.design`.
	 * Tests inject a static, offline resolver.
	 */
	resolveIcon?: IconResolver
	/**
	 * Writer factory seam: returns a fresh writer instance per call. Defaults to
	 * `() => new TsPptx()`. Lets the emit layer be tested
	 * against an injected/mock instance.
	 */
	pptxFactory?: () => unknown
	/** Delivery mode. Defaults to `'download'`. Replaces the old `__TEST__` hack. */
	output?: OutputMode
	/**
	 * Vectorize SVG/icon graphics into editable PowerPoint freeform shapes
	 * (`custGeom`) instead of rasterizing them to PNG. Opt-in: anything that
	 * can't be faithfully vectorized (transforms, gradient paints,
	 * `<use>`/`<text>`, missing viewBox) silently falls back to the raster image,
	 * so enabling this never drops content.
	 */
	vectorizeSvg?: boolean
}

export interface ConvertResult {
	warnings: Warning[]
	slideCount: number
	totalSlideCount: number
	stopped: boolean
	/** Present when `output: 'base64'`. */
	base64?: string
	/** Present when `output: 'blob'`. */
	blob?: Blob
	/** Present when `output: 'pptx-instance'`: the underlying ts-pptx instance. */
	pptx?: unknown
}

/**
 * Convert a full HTML document (head + slide sections) into PPTX.
 *
 * Delegates to the lifted engine (on `@shbernal/ts-pptx`). Delivery is
 * controlled by `opts.output` (default `'download'`); `'base64'` / `'blob'` /
 * `'pptx-instance'` return the deck on the result instead. Icons resolve through
 * `opts.resolveIcon`, and the writer through `opts.pptxFactory`, both defaulted.
 */
export async function convertDeck(fullHtmlString: string, opts?: ConvertOptions): Promise<ConvertResult> {
	return convertDeckEngine(fullHtmlString, opts) as Promise<ConvertResult>
}

/**
 * Convert a single slide's HTML into its IR slide model (plus warnings).
 * Per-slide entry point for testing and granularity: it renders the slide and
 * returns the IR (the `extract` layer's output) without emitting PPTX, so the
 * model boundary can be asserted directly. Icons resolve through `opts.resolveIcon`.
 */
export async function convertSlide(
	headHTML: string,
	slideHTML: string,
	opts?: ConvertOptions
): Promise<{ model: SlideModel; warnings: Warning[] }> {
	return convertSlideEngine(headHTML, slideHTML, opts) as Promise<{ model: SlideModel; warnings: Warning[] }>
}
