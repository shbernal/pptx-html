/**
 * The heuristic lane's orchestrator — parse the deck, render each slide in a
 * hidden frame, read its model, emit it onto a ts-pptx slide, then deliver the
 * package.
 *
 * This is the lane for HTML that carries no IR island: there is no source
 * package to import, so the model is *inferred* from what the browser painted
 * and the output is best-effort by construction. It shares no code with the
 * exact lane (`src/import` → `src/render` → `src/parse` → `src/emit`) beyond the
 * writer itself, which is deliberate — merging them would quietly weaken the
 * lane that carries a guarantee.
 *
 * The seams, all of them options on the `opts` argument:
 *  - the writer is `@shbernal/ts-pptx`, imported here and overridable via
 *    `opts.pptxFactory`;
 *  - delivery is the `opts.output` option (`download` | `base64` | `blob` |
 *    `pptx-instance`);
 *  - icons resolve through `opts.resolveIcon` (default: `defaultResolveIcon`).
 *
 * The entry points are ES module exports, re-exported from `index.ts`.
 */

import { ShapeType, TsPptx } from '@shbernal/ts-pptx'
import { bytesOfBase64 } from '../base64'
import { DEFAULT_FONT } from '../constants'
import { EMU_PER_INCH } from '../ir/render'
import { createHiddenFrame, settleFrame } from './frame'
import { defaultResolveIcon, type IconResolver, inlineDeckIcons } from './icons'
import type { SlideModel, SlideSize } from './model'
import { composeSlideDocument, parseDeckHtml } from './parse'
import type { PptxDeck } from './pptx-types'
import { readSlideModel } from './read'
import { addModelToSlide } from './slide'
import { vectorizeSvgImages } from './svg'

/** How a non-fatal issue surfaced during conversion. */
export interface Warning {
	slide?: number
	code?: string
	message: string
}

/**
 * Progress, as this lane actually reports it: once per icon-resolution pass,
 * once per slide as it starts, and once at the end. (An earlier declaration
 * listed `parse`/`render`/`extract`/`emit` phases that were never emitted — a
 * type nothing could match on, which is worse than none.)
 */
export interface ProgressEvent {
	phase: 'icons' | 'slide' | 'finalize'
	/** Zero-based slide index, on `phase: 'slide'`. */
	index?: number
	completed?: number
	total?: number
	stopped?: boolean
}

/** Where the converted deck should be delivered. */
export type OutputMode = 'download' | 'base64' | 'blob' | 'pptx-instance'

export interface ConvertOptions {
	author?: string
	title?: string
	lang?: string
	/** Used for `output: 'download'`. */
	fileName?: string
	fontConfig?: Record<string, string>
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
	 * `() => new TsPptx()`. Lets the emit layer be tested against an
	 * injected/mock instance.
	 */
	pptxFactory?: () => unknown
	/** Delivery mode. Defaults to `'download'`. */
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

function getPptSize(pptx: PptxDeck): SlideSize {
	return {
		width: pptx.presLayout ? pptx.presLayout.width / EMU_PER_INCH : 10,
		height: pptx.presLayout ? pptx.presLayout.height / EMU_PER_INCH : 5.625,
	}
}

function collectWarnings(model: SlideModel, size: SlideSize): string[] {
	const warnings: string[] = []
	const tolerance = 0.03
	for (const item of model.items) {
		const p = item.position
		if (!p) continue
		const over: string[] = []
		if (p.x < -tolerance) over.push('left')
		if (p.y < -tolerance) over.push('top')
		if (p.x + p.w > size.width + tolerance) over.push('right')
		if (p.y + p.h > size.height + tolerance) over.push('bottom')
		if (over.length && item.kind !== 'decor') {
			warnings.push(
				'NEEDS REVIEW: "' + (item.label || item.type) + '" extends beyond the slide on the ' + over.join('/') + ' edge.'
			)
		}
	}
	return warnings
}

const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'

/**
 * A `.pptx` blob from byte chunks, with the cast the DOM types force in one
 * place. `BlobPart` wants an `ArrayBufferView<ArrayBuffer>`, and `Uint8Array` is
 * declared over `ArrayBufferLike`, which also admits a `SharedArrayBuffer` no
 * writer here returns. The constructor takes either at run time.
 */
function pptxBlobOf(chunks: Uint8Array[]): Blob {
	return new Blob(chunks as BlobPart[], { type: PPTX_MIME })
}

function base64ToPptxBlob(base64: string): Blob {
	return pptxBlobOf([bytesOfBase64(base64)])
}

function downloadBlob(blob: Blob, fileName?: string): void {
	const url = URL.createObjectURL(blob)
	const a = document.createElement('a')
	a.href = url
	a.download = fileName || 'deck.pptx'
	document.body.appendChild(a)
	a.click()
	setTimeout(() => {
		try {
			a.remove()
		} catch {
			/* already detached */
		}
		URL.revokeObjectURL(url)
	}, 0)
}

// Create a fresh writer instance: the bundled ts-pptx by default, or whatever
// `opts.pptxFactory` returns (a mock, or a differently-configured writer).
function createPptx(opts: ConvertOptions): PptxDeck {
	if (typeof opts.pptxFactory === 'function') return opts.pptxFactory() as PptxDeck
	return new TsPptx() as unknown as PptxDeck
}

// Render one slide section in a hidden frame and read its model. Shared by the
// deck loop (buildDeck) and the per-slide entry point (convertSlide) so both go
// through the identical extract path. Returns the model plus the warnings raised
// on the way — the extractor boundary's, then the layout (off-slide) ones; the
// caller owns slide numbering and emit.
async function renderSlideModel(
	headHTML: string,
	slideHTML: string,
	lang: string,
	size: SlideSize,
	opts: ConvertOptions
): Promise<{ model: SlideModel; warnings: string[] }> {
	const frame = createHiddenFrame(composeSlideDocument(headHTML, slideHTML, lang))
	try {
		await frame.load()
		await settleFrame(frame)
		const doc = frame.iframe.contentDocument
		if (!doc) throw new Error('Slide render frame was closed before it could be measured.')
		const root = doc.querySelector('.slide') || doc.body
		const rect = root.getBoundingClientRect()
		frame.resize(rect.width || 1280, rect.height || 720)
		await settleFrame(frame)
		const read = await readSlideModel(frame, size, { ...DEFAULT_FONT, ...(opts.fontConfig || {}) })
		// Opt-in vectorization. Replace SVG/icon raster images
		// with editable custGeom paths; anything unsupported stays a raster image.
		if (opts.vectorizeSvg) vectorizeSvgImages(read.model)
		return { model: read.model, warnings: [...read.warnings, ...collectWarnings(read.model, size)] }
	} finally {
		frame.close()
	}
}

async function buildDeck(headHTML: string, slides: string[], lang: string, opts: ConvertOptions) {
	const pptx = createPptx(opts)
	pptx.layout = 'LAYOUT_16x9'
	pptx.author = opts.author || 'pptx-html'
	pptx.subject = opts.title || 'pptx-html PPTX export'
	pptx.company = 'pptx-html'
	// No deck-level language is set here, and there is nothing to set: the writer
	// declares `layout`/`author`/`company`/`subject`/`theme` and no `lang`, and
	// `ThemeProps` has none either. Both spellings used to be assigned anyway —
	// they landed as own properties on the instance and were never written into the
	// package. `lang` is a *run* option, so the language reaches the deck the only
	// way the writer can carry it: on every `addText` call, from `addModelToSlide`.
	pptx.theme = { headFontFace: DEFAULT_FONT.latin, bodyFontFace: DEFAULT_FONT.latin }
	const size = getPptSize(pptx)
	const warnings: Warning[] = []
	const shouldStop = () => !!opts.shouldStop?.()
	let completedSlides = 0
	let stopped = false
	// Resolve icons once and inline them, so the per-slide frames never depend on
	// Iconify custom-element timing; falls back to raw slides if all resolvers fail.
	opts.onProgress?.({ phase: 'icons', total: slides.length })
	const inlined = await inlineDeckIcons(headHTML, slides, opts.resolveIcon || defaultResolveIcon)
	if (inlined) slides = inlined
	// `entries()` rather than an index loop: the slide HTML comes out of the
	// iterator already known to exist, which is the same guarantee the bounds check
	// gave and one the compiler can see.
	for (const [index, slideHtml] of slides.entries()) {
		if (shouldStop()) {
			stopped = true
			break
		}
		opts.onProgress?.({ phase: 'slide', index, total: slides.length })
		try {
			const rendered = await renderSlideModel(headHTML, slideHtml, lang, size, opts)
			warnings.push(...rendered.warnings.map((message) => ({ slide: index + 1, message })))
			const slide = pptx.addSlide()
			completedSlides++
			// ts-pptx exposes `ShapeType` as a module export, not off the instance, so
			// the emit layer's writer seam is supplied here rather than being read off
			// `pptx` (which may be a mock from `opts.pptxFactory`).
			const issues = await addModelToSlide({ ShapeType }, slide, rendered.model, size, lang || 'en')
			warnings.push(...issues.map((message) => ({ slide: index + 1, message })))
		} catch (error) {
			warnings.push({
				slide: index + 1,
				message: 'SLIDE EXPORT ISSUE: ' + (error instanceof Error && error.message ? error.message : String(error)),
			})
		}
		if (shouldStop()) {
			stopped = true
			break
		}
	}
	return { pptx, warnings, stopped, completedSlides, totalSlides: slides.length }
}

export async function convertDeck(fullHtmlString: string, opts: ConvertOptions = {}): Promise<ConvertResult> {
	const parsed = parseDeckHtml(fullHtmlString)
	if (!parsed.slides.length) throw new Error('Export input contains no slide sections.')
	const built = await buildDeck(parsed.headHTML, parsed.slides, parsed.lang, opts)
	if (built.stopped && !built.completedSlides) throw new Error('Export stopped before any slide was created.')
	opts.onProgress?.({
		phase: 'finalize',
		total: parsed.slides.length,
		stopped: built.stopped,
		completed: built.completedSlides,
	})

	const result: ConvertResult = {
		warnings: built.warnings,
		slideCount: built.stopped ? built.completedSlides : parsed.slides.length,
		stopped: built.stopped,
		totalSlideCount: parsed.slides.length,
	}

	return deliverDeck(built.pptx, result, opts)
}

// Deliver per `opts.output` (default 'download').
// Honors `output: 'download' | 'base64' | 'blob' | 'pptx-instance'`.
async function deliverDeck(pptx: PptxDeck, result: ConvertResult, opts: ConvertOptions): Promise<ConvertResult> {
	const output = opts.output || 'download'
	if (output === 'pptx-instance') {
		// Hand back the live writer, unserialized.
		result.pptx = pptx
		return result
	}

	// Written and delivered as-is. There used to be a post-write OOXML repair pass
	// here; see `test/oracle/writer-output.test.ts` for why there is not.
	//
	// Two of the three deliveries want bytes, and `toBytes()` hands them over
	// directly. Going through base64 for those meant encoding a multi-megabyte
	// archive with `btoa` and decoding it again with a per-character loop, in the
	// browser, to arrive back where the writer started. Only `output: 'base64'`
	// asks for the encoded form, so only it pays for one.
	if (output === 'base64') {
		result.base64 = await pptx.write({ outputType: 'base64' })
		return result
	}
	const blob = await pptxBlob(pptx)
	if (output === 'blob') {
		result.blob = blob
	} else {
		// 'download': browser-only delivery, identical to the original behavior.
		downloadBlob(blob, opts.fileName || 'deck.pptx')
	}
	return result
}

/**
 * The deck as a `Blob`, straight from `toBytes()` where the writer has it.
 *
 * The `write` arm is what keeps `opts.pptxFactory` open to a mock that predates
 * ts-pptx 3.3.0, and it is the only reason the base64 decoder is still reachable
 * from this lane.
 */
async function pptxBlob(pptx: PptxDeck): Promise<Blob> {
	if (pptx.toBytes) return pptxBlobOf([await pptx.toBytes()])
	return base64ToPptxBlob(await pptx.write({ outputType: 'base64' }))
}

/**
 * Per-slide entry point: render a single slide's HTML and return its model (plus
 * warnings), without emitting PPTX. This is the lane's model boundary made
 * directly observable — the browser test harness asserts on the model here, and
 * callers can use it for granular conversion. Icons resolve through the same
 * `opts.resolveIcon` seam, and the slide is sized for the 16:9 layout.
 */
export async function convertSlide(
	headHTML: string,
	slideHTML: string,
	opts: ConvertOptions = {}
): Promise<{ model: SlideModel; warnings: Warning[] }> {
	const pptx = createPptx(opts)
	pptx.layout = 'LAYOUT_16x9'
	const size = getPptSize(pptx)
	const inlined = await inlineDeckIcons(headHTML, [slideHTML], opts.resolveIcon || defaultResolveIcon)
	const slide = inlined?.[0] || slideHTML
	const rendered = await renderSlideModel(headHTML, slide, opts.lang || 'en', size, opts)
	return { model: rendered.model, warnings: rendered.warnings.map((message) => ({ message })) }
}

/** Version tag for the engine. */
export const ENGINE_VERSION = 'quality-browser-v1'
