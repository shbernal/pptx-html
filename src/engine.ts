// @ts-nocheck
/**
 * Deck orchestrator — ties the extract, emit and repair layers together: parse
 * the deck, render each slide in a hidden frame, read its IR model, emit it onto
 * a ts-pptx slide, then repair and deliver the package.
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
import html2canvas from 'html2canvas'
import { DEFAULT_FONT, EMU_PER_IN } from './constants'
import { createHiddenFrame, settleFrame } from './extract/frame'
import { defaultResolveIcon, inlineDeckIcons } from './extract/icons'
import { composeSlideDocument, parseDeckHtml } from './extract/parse'
import { readSlideModel } from './extract/read'
import { vectorizeSvgImages } from './extract/svg'
import { addModelToSlide } from './emit/slide'
import { repairPptxBase64 } from './repair/repair'

function getPptSize(pptx) {
	return {
		width: pptx.presLayout ? pptx.presLayout.width / EMU_PER_IN : 10,
		height: pptx.presLayout ? pptx.presLayout.height / EMU_PER_IN : 5.625
	}
}

function collectWarnings(model, size) {
	const warnings = []
	const tolerance = 0.03
	for (const item of model.items) {
		const p = item.position
		if (!p) continue
		const over = []
		if (p.x < -tolerance) over.push('left')
		if (p.y < -tolerance) over.push('top')
		if (p.x + p.w > size.width + tolerance) over.push('right')
		if (p.y + p.h > size.height + tolerance) over.push('bottom')
		if (over.length && item.kind !== 'decor') {
			warnings.push('NEEDS REVIEW: "' + (item.label || item.type) + '" extends beyond the slide on the ' + over.join('/') + ' edge.')
		}
	}
	return warnings
}

const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'

function base64ToPptxBlob(base64) {
	const bytes = atob(base64)
	const chunks = []
	for (let i = 0; i < bytes.length; i += 32768) {
		const slice = bytes.slice(i, i + 32768)
		const arr = new Uint8Array(slice.length)
		for (let j = 0; j < slice.length; j++) arr[j] = slice.charCodeAt(j)
		chunks.push(arr)
	}
	return new Blob(chunks, { type: PPTX_MIME })
}

function downloadBase64Pptx(base64, fileName) {
	const blob = base64ToPptxBlob(base64)
	const url = URL.createObjectURL(blob)
	const a = document.createElement('a')
	a.href = url
	a.download = fileName || 'deck.pptx'
	document.body.appendChild(a)
	a.click()
	setTimeout(() => {
		try { a.remove() } catch (e) { /* already detached */ }
		URL.revokeObjectURL(url)
	}, 0)
}

// Create a fresh writer instance: the bundled ts-pptx by default, or whatever
// `opts.pptxFactory` returns (a mock, or a differently-configured writer).
function createPptx(opts) {
	if (typeof opts.pptxFactory === 'function') return opts.pptxFactory()
	return new TsPptx()
}

// Render one slide section in a hidden frame and read its IR model. Shared by the
// deck loop (buildDeck) and the per-slide entry point (convertSlide) so both go
// through the identical extract path. Returns the IR model plus the layout
// (off-slide) warnings; the caller owns slide numbering and emit.
async function renderSlideModel(headHTML, slideHTML, lang, size, opts) {
	const frame = createHiddenFrame(composeSlideDocument(headHTML, slideHTML, lang))
	try {
		await frame.load()
		await settleFrame(frame)
		const root = frame.iframe.contentDocument.querySelector('.slide') || frame.iframe.contentDocument.body
		const rect = root.getBoundingClientRect()
		frame.resize(rect.width || 1280, rect.height || 720)
		await settleFrame(frame)
		const model = await readSlideModel(frame, size, { ...DEFAULT_FONT, ...(opts.fontConfig || {}) })
		// Opt-in vectorization. Replace SVG/icon raster images
		// with editable custGeom paths; anything unsupported stays a raster image.
		if (opts.vectorizeSvg) vectorizeSvgImages(model)
		return { model, warnings: collectWarnings(model, size) }
	} finally {
		frame.close()
	}
}

async function buildDeck(headHTML, slides, lang, opts) {
	const pptx = createPptx(opts)
	pptx.layout = 'LAYOUT_16x9'
	pptx.author = opts.author || 'dom2pptx'
	pptx.subject = opts.title || 'dom2pptx PPTX export'
	pptx.company = 'dom2pptx'
	pptx.lang = lang || 'en'
	pptx.theme = {
		headFontFace: DEFAULT_FONT.latin,
		bodyFontFace: DEFAULT_FONT.latin,
		lang: lang || 'en'
	}
	const size = getPptSize(pptx)
	const warnings = []
	const shouldStop = () => !!(opts.shouldStop && opts.shouldStop())
	let completedSlides = 0
	let stopped = false
	// Resolve icons once and inline them, so the per-slide frames never depend on
	// Iconify custom-element timing; falls back to raw slides if all resolvers fail.
	if (opts.onProgress) opts.onProgress({ phase: 'icons', total: slides.length })
	const inlined = await inlineDeckIcons(headHTML, slides, opts.resolveIcon || defaultResolveIcon)
	if (inlined) slides = inlined
	for (let index = 0; index < slides.length; index++) {
		if (shouldStop()) {
			stopped = true
			break
		}
		if (opts.onProgress) opts.onProgress({ phase: 'slide', index, total: slides.length })
		try {
			const rendered = await renderSlideModel(headHTML, slides[index], lang, size, opts)
			warnings.push(...rendered.warnings.map((message) => ({ slide: index + 1, message })))
			const slide = pptx.addSlide()
			completedSlides++
			// ts-pptx exposes `ShapeType` as a module export, not off the instance, so
			// the emit layer's writer seam is supplied here rather than being read off
			// `pptx` (which may be a mock from `opts.pptxFactory`).
			const issues = await addModelToSlide({ ShapeType }, slide, rendered.model, size)
			warnings.push(...(issues || []).map((message) => ({ slide: index + 1, message })))
		} catch (error) {
			warnings.push({ slide: index + 1, message: 'SLIDE EXPORT ISSUE: ' + (error && error.message ? error.message : String(error)) })
		}
		if (shouldStop()) {
			stopped = true
			break
		}
	}
	return { pptx, warnings, stopped, completedSlides, totalSlides: slides.length }
}

export async function convertDeck(fullHtmlString, opts) {
	opts = opts || {}
	const parsed = parseDeckHtml(fullHtmlString)
	if (!parsed.slides.length) throw new Error('Export input contains no slide sections.')
	const built = await buildDeck(parsed.headHTML, parsed.slides, parsed.lang, opts)
	if (built.stopped && !built.completedSlides) throw new Error('Export stopped before any slide was created.')
	if (opts.onProgress) opts.onProgress({ phase: 'finalize', total: parsed.slides.length, stopped: built.stopped, completed: built.completedSlides })

	const result = {
		warnings: built.warnings,
		slideCount: built.stopped ? built.completedSlides : parsed.slides.length,
		stopped: built.stopped,
		totalSlideCount: parsed.slides.length
	}

	return deliverDeck(built.pptx, result, opts)
}

// Deliver per `opts.output` (default 'download'), replacing the old __TEST__ hack.
// Shared by the vector (`convertDeck`) and raster (`convertDeckRaster`) paths so
// both honor `output: 'download' | 'base64' | 'blob' | 'pptx-instance'` identically.
async function deliverDeck(pptx, result, opts) {
	const output = opts.output || 'download'
	if (output === 'pptx-instance') {
		// Hand back the live writer. Post-write OOXML repairs operate on serialized
		// bytes, so they cannot apply here — the caller owns write()/repair if wanted.
		result.pptx = pptx
		return result
	}

	const base64 = await repairPptxBase64(await pptx.write({ outputType: 'base64' }))
	if (output === 'base64') {
		result.base64 = base64
	} else if (output === 'blob') {
		result.blob = base64ToPptxBlob(base64)
	} else {
		// 'download': browser-only delivery, identical to the original behavior.
		downloadBase64Pptx(base64, opts.fileName || 'deck.pptx')
	}
	return result
}

// Render one slide section in a hidden frame and rasterize it with html2canvas.
// Capture the rendered `.slide` at 2x and return a JPEG data URL sized for the
// 16:9 layout.
async function rasterizeSlide(headHTML, slideHTML, lang) {
	const frame = createHiddenFrame(composeSlideDocument(headHTML, slideHTML, lang))
	try {
		await frame.load()
		await settleFrame(frame)
		const doc = frame.iframe.contentDocument
		const target = doc.querySelector('.slide') || doc.body
		const canvas = await html2canvas(target, { scale: 2, useCORS: true, backgroundColor: '#fff', width: 1280, height: 720 })
		return canvas.toDataURL('image/jpeg', 0.92)
	} finally {
		frame.close()
	}
}

/**
 * Image-based fallback: rasterize each slide with html2canvas and place it as a
 * full-bleed picture. Use this when the editable (vector) `convertDeck` path
 * fails — a file is still produced, just not editable. Keeping it here means
 * consumers never import `@shbernal/ts-pptx`/`html2canvas` directly (this package
 * owns both).
 *
 * Signature and `ConvertResult` shape mirror `convertDeck`; delivery is the same
 * `opts.output` switch. `onProgress` emits `{ phase: 'slide', index, total }` and
 * `{ phase: 'finalize', ... }`, matching the vector path for consumer parity.
 */
export async function convertDeckRaster(fullHtmlString, opts) {
	opts = opts || {}
	const parsed = parseDeckHtml(fullHtmlString)
	if (!parsed.slides.length) throw new Error('Export input contains no slide sections.')

	const pptx = createPptx(opts)
	pptx.layout = 'LAYOUT_16x9'
	pptx.author = opts.author || 'dom2pptx'
	pptx.subject = opts.title || 'dom2pptx PPTX export'
	pptx.company = 'dom2pptx'
	pptx.lang = parsed.lang || 'en'
	const size = getPptSize(pptx)

	const warnings = []
	const shouldStop = () => !!(opts.shouldStop && opts.shouldStop())
	let completedSlides = 0
	let stopped = false
	for (let index = 0; index < parsed.slides.length; index++) {
		if (shouldStop()) { stopped = true; break }
		if (opts.onProgress) opts.onProgress({ phase: 'slide', index, total: parsed.slides.length })
		try {
			const data = await rasterizeSlide(parsed.headHTML, parsed.slides[index], parsed.lang)
			pptx.addSlide().addImage({ data, x: 0, y: 0, w: size.width, h: size.height })
			completedSlides++
		} catch (error) {
			warnings.push({ slide: index + 1, message: 'SLIDE EXPORT ISSUE: ' + (error && error.message ? error.message : String(error)) })
		}
		if (shouldStop()) { stopped = true; break }
	}
	if (stopped && !completedSlides) throw new Error('Export stopped before any slide was created.')
	if (opts.onProgress) opts.onProgress({ phase: 'finalize', total: parsed.slides.length, stopped, completed: completedSlides })

	const result = {
		warnings,
		slideCount: stopped ? completedSlides : parsed.slides.length,
		stopped,
		totalSlideCount: parsed.slides.length
	}
	return deliverDeck(pptx, result, opts)
}

/**
 * Per-slide entry point: render a single slide's HTML and return its IR model
 * (plus warnings), without emitting PPTX. This is the IR boundary made directly
 * observable — the browser test harness asserts on the model here, and callers
 * can use it for granular conversion. Icons resolve through the same
 * `opts.resolveIcon` seam, and the slide is sized for the 16:9 layout.
 */
export async function convertSlide(headHTML, slideHTML, opts) {
	opts = opts || {}
	const pptx = createPptx(opts)
	pptx.layout = 'LAYOUT_16x9'
	const size = getPptSize(pptx)
	const inlined = await inlineDeckIcons(headHTML, [slideHTML], opts.resolveIcon || defaultResolveIcon)
	const slide = (inlined && inlined[0]) || slideHTML
	const rendered = await renderSlideModel(headHTML, slide, opts.lang || 'en', size, opts)
	return { model: rendered.model, warnings: rendered.warnings.map((message) => ({ message })) }
}

/** Version tag for the engine. */
export const ENGINE_VERSION = 'quality-browser-v1'
