/**
 * Shared helpers for the browser test layer: an offline icon
 * resolver, a ts-pptx `read` loader, and a position-free IR summariser for stable
 * structural snapshots.
 */

import { Presentation } from '@shbernal/ts-pptx/read'

/**
 * Static, offline icon resolver injected via `opts.resolveIcon`. Returns a fixed
 * filled square for any icon name so icon tests are deterministic and never touch
 * the network (the default resolver fetches from api.iconify.design).
 */
export function staticResolveIcon(_name: string): Promise<string | null> {
	return Promise.resolve(
		'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M4 4h16v16H4z"/></svg>'
	)
}

/** Decode a base64 PPTX into bytes ts-pptx's `read` model can load (no Buffer in the browser). */
export function base64ToBytes(base64: string): Uint8Array {
	const binary = atob(base64)
	const bytes = new Uint8Array(binary.length)
	for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
	return bytes
}

interface ReadShape {
	text?: string
}
interface ReadSlide {
	shapes: ReadShape[]
}
export interface ReadDeck {
	slides: ReadSlide[]
}

/** Parse a base64 deck back with ts-pptx's `read` Presentation model. */
export function loadDeck(base64: string): Promise<ReadDeck> {
	return Presentation.load(base64ToBytes(base64))
}

/** Concatenate the text of all shapes across every slide. */
export function allDeckText(deck: ReadDeck): string {
	return deck.slides.flatMap((slide) => slide.shapes.map((shape) => shape.text || '')).join(' ')
}

interface RawRun {
	text: string
}
interface RawItem {
	type: string
	z?: number
	tag?: string
	text?: string | RawRun[]
	style?: { color?: string; bold?: boolean }
	rows?: RawRun[][]
	src?: string
	color?: string
	fill?: string
}
export interface RawModel {
	background: unknown
	items: RawItem[]
}

/** Flatten an IR text value (plain string or run array) to its text content. */
export function runText(text: string | RawRun[] | undefined): string {
	if (typeof text === 'string') return text
	if (Array.isArray(text)) return text.map((run) => run.text).join('')
	return ''
}

function srcKind(src: string | undefined): string {
	if (/^data:image\/svg/i.test(src || '')) return 'svg'
	if (/^data:image\/png/i.test(src || '')) return 'png'
	if (/^data:/i.test(src || '')) return 'data'
	return 'url'
}

/**
 * Position-free structural summary of an IR model. Pixel positions depend on
 * font metrics and Chromium layout, so they are dropped from the snapshot;
 * everything that defines the *structure* (types, z-order, text, colours, the
 * kind of each image source) is kept. Snapshot this, not the raw model.
 */
export function summarizeModel(model: RawModel) {
	return {
		background: model.background,
		items: model.items.map((item) => {
			const base: Record<string, unknown> = { type: item.type, z: item.z }
			if (item.type === 'text') {
				base.tag = item.tag
				base.text = runText(item.text)
				base.color = item.style?.color
				base.bold = !!item.style?.bold
			} else if (item.type === 'list') {
				base.runs = Array.isArray(item.text) ? item.text.map((run) => run.text) : []
			} else if (item.type === 'table') {
				base.rows = (item.rows || []).map((row) => row.map((cell) => cell.text))
			} else if (item.type === 'image') {
				base.src = srcKind(item.src)
			} else if (item.type === 'line') {
				base.color = item.color
			} else if (item.type === 'shape') {
				base.fill = item.fill
			}
			return base
		}),
	}
}
