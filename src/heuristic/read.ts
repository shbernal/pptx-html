/**
 * The boundary the browser extractor is quarantined behind.
 *
 * `extractor.ts` is stringified and `eval`'d inside the slide iframe, so it
 * cannot be type-checked where it runs: its realm has no modules and its return
 * value crosses back as `unknown`. Everything on this side of that call *is*
 * checked, and this file is what makes the two meet — it validates the extracted
 * model against `./model.ts` before any of it reaches the emitter.
 *
 * That is the whole argument for the file. A `@ts-nocheck` on the extractor is
 * tolerable only while nothing downstream trusts its output on faith; the moment
 * a malformed item can reach the writer unchecked, the suppression has spread
 * past the one function that earns it.
 *
 * **A rejected item is warned about, never silently dropped.** The lane's
 * promise is the charter's — modeled, carried or warned — and a missing shape
 * that nobody mentioned is the one outcome it rules out. A malformed *deck*, by
 * contrast, throws: an extractor that returns no `items` array is a bug in this
 * package, not a property of the page, and continuing with an empty slide would
 * report that bug as a blank deck.
 *
 * The `win.eval(fn.toString())` indirection is intentional: the extractor must
 * run against the iframe's `document` / `getComputedStyle`, not the host's.
 */

import type { SlideFrame } from './frame'
import type { Background, Item, Rect, SlideModel, SlideSize } from './model'
import { browserExtractor } from './extractor'

/** Font-family mapping handed to the extractor; shape of `DEFAULT_FONT`. */
export type FontConfig = Record<string, string>

/** An extracted slide plus whatever the boundary refused on the way through. */
export interface ExtractedSlide {
	model: SlideModel
	warnings: string[]
}

/** The extractor's own parameter object, as it is called inside the frame. */
interface ExtractorConfig {
	slideWidthIn: number
	slideHeightIn: number
	fonts: FontConfig
}

/** The iframe realm's `eval`, which is all this file needs from a window. */
type FrameRealm = { eval: (source: string) => unknown }

const ITEM_TYPES = new Set(['text', 'list', 'shape', 'image', 'line', 'table', 'path'])

function isRecord(value: unknown): value is Record<string, unknown> {
	return typeof value === 'object' && value !== null
}

function isFiniteNumber(value: unknown): value is number {
	return typeof value === 'number' && Number.isFinite(value)
}

function rectOf(value: unknown): Rect | null {
	if (!isRecord(value)) return null
	const { x, y, w, h } = value
	if (!isFiniteNumber(x) || !isFiniteNumber(y) || !isFiniteNumber(w) || !isFiniteNumber(h)) return null
	return { x, y, w, h }
}

/**
 * A background the emitter can spend. An unrecognised one is not fatal — a slide
 * with the wrong backdrop still carries its content — so it degrades to white
 * and says so.
 */
function backgroundOf(value: unknown, warnings: string[]): Background {
	if (isRecord(value)) {
		if (value.type === 'color' && typeof value.value === 'string') return { type: 'color', value: value.value }
		if (value.type === 'image' && typeof value.src === 'string') {
			const fallback = typeof value.fallback === 'string' ? value.fallback : undefined
			return fallback === undefined ? { type: 'image', src: value.src } : { type: 'image', src: value.src, fallback }
		}
	}
	if (value !== undefined) warnings.push('NEEDS REVIEW: the slide background could not be read and was left white.')
	return { type: 'color', value: 'FFFFFF' }
}

/**
 * What each item type needs beyond a position before the emitter will do
 * anything useful with it. Checked here rather than in `slide.ts`, so a
 * structurally broken item is named once at the boundary instead of surfacing as
 * a writer exception with no idea which element it came from.
 */
function payloadComplaint(item: Record<string, unknown>): string | null {
	switch (item.type) {
		case 'text':
		case 'list':
			return typeof item.text === 'string' || Array.isArray(item.text) ? null : 'has no text'
		case 'image':
			return typeof item.src === 'string' && item.src !== '' ? null : 'has no source'
		case 'table':
			return Array.isArray(item.rows) && item.rows.every(Array.isArray) ? null : 'has no rows'
		case 'path':
			return Array.isArray(item.points) && item.points.length > 0 ? null : 'has no path points'
		default:
			return null
	}
}

function itemOf(value: unknown, index: number, warnings: string[]): Item | null {
	const label = `item ${index + 1}`
	if (!isRecord(value) || typeof value.type !== 'string' || !ITEM_TYPES.has(value.type)) {
		warnings.push(`NEEDS REVIEW: ${label} has an unrecognised type and was left out of the slide.`)
		return null
	}
	const named = typeof value.label === 'string' ? `"${value.label}"` : `a ${value.type}`
	const position = rectOf(value.position)
	if (!position) {
		warnings.push(`NEEDS REVIEW: ${named} had no usable position and was left out of the slide.`)
		return null
	}
	const complaint = payloadComplaint(value)
	if (complaint) {
		warnings.push(`NEEDS REVIEW: ${named} ${complaint} and was left out of the slide.`)
		return null
	}
	// Everything past the discriminant, the position and the payload check is
	// passed through as the extractor produced it: the option bags are opaque by
	// design (see `./model.ts`), so validating them would mean inventing a schema
	// the writer does not publish.
	return { ...value, type: value.type, position, z: isFiniteNumber(value.z) ? value.z : 0 } as Item
}

/**
 * Validate what came back out of the frame. Exported for the boundary's own
 * tests, which is the only way to exercise it without a browser.
 */
export function coerceSlideModel(value: unknown): ExtractedSlide {
	if (!isRecord(value) || !Array.isArray(value.items)) {
		throw new Error('The slide extractor returned no model; the slide could not be read.')
	}
	const warnings: string[] = []
	const background = backgroundOf(value.background, warnings)
	const items: Item[] = []
	for (const [index, entry] of value.items.entries()) {
		const item = itemOf(entry, index, warnings)
		if (item) items.push(item)
	}
	const model: SlideModel = { background, items }
	if (typeof value.notes === 'string' && value.notes !== '') model.notes = value.notes
	return { model, warnings }
}

export async function readSlideModel(
	frame: SlideFrame,
	slideSize: SlideSize,
	fontConfig: FontConfig
): Promise<ExtractedSlide> {
	const win = frame.iframe.contentWindow as FrameRealm | null
	if (!win) throw new Error('Slide render frame was closed before its model could be read.')
	const run = win.eval(`(${browserExtractor.toString()})`) as (config: ExtractorConfig) => Promise<unknown>
	return coerceSlideModel(
		await run({ slideWidthIn: slideSize.width, slideHeightIn: slideSize.height, fonts: fontConfig })
	)
}
