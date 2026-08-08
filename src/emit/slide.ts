/**
 * Emit layer — IR slide model → ts-pptx objects.
 *
 * Pure and isomorphic (no DOM beyond the SVG→PNG raster helper, which is guarded
 * by a browser `Image`/`canvas`): drives a ts-pptx writer through `addImage` /
 * `addShape` / `addTable` / `addText`. Items are typed against the IR
 * (`../ir/model`); the writer/slide are the structural shapes in `./pptx-types`,
 * so this layer stays decoupled from the full writer types and unit-testable
 * against an injected writer.
 */

import { DEFAULT_FONT, PX_PER_IN } from '../constants'
import type { Item, Rect, SlideModel, SlideSize, TextItem } from '../ir/model'
import { emitPathItem } from './custgeom'
import type { PptxSlide, PptxWriter } from './pptx-types'

async function svgDataToPng(dataUrl: string, widthPx: number, heightPx: number): Promise<string> {
	const img = await new Promise<HTMLImageElement>((resolve, reject) => {
		const node = new Image()
		node.onload = () => resolve(node)
		node.onerror = () => reject(new Error('SVG graphic could not be prepared for PPTX export.'))
		node.src = dataUrl
	})
	const canvas = document.createElement('canvas')
	canvas.width = Math.max(1, Math.round(widthPx || img.naturalWidth || img.width || 1))
	canvas.height = Math.max(1, Math.round(heightPx || img.naturalHeight || img.height || 1))
	canvas.getContext('2d')!.drawImage(img, 0, 0, canvas.width, canvas.height)
	return canvas.toDataURL('image/png')
}

function imageRef(src: string): { data: string } | { path: string } {
	if (typeof src === 'string' && /^data:/i.test(src)) return { data: src }
	return { path: src }
}

function sortedItems(items: Item[]): Item[] {
	return items
		.map((item, index) => ({ item, index, z: Number.isFinite(item.z) ? item.z : 0 }))
		.sort((a, b) => a.z - b.z || a.index - b.index)
		.map((entry) => entry.item)
}

export async function addModelToSlide(pptx: PptxWriter, slide: PptxSlide, model: SlideModel, slideSize: SlideSize): Promise<string[]> {
	const issues: string[] = []
	if (model.background.type === 'color') {
		slide.background = { color: model.background.value || 'FFFFFF' }
	} else if (model.background.type === 'image' && model.background.src) {
		try {
			slide.background = { color: model.background.fallback || 'FFFFFF' }
			slide.addImage({ ...imageRef(model.background.src), x: 0, y: 0, w: slideSize.width, h: slideSize.height, sizing: { type: 'cover', w: slideSize.width, h: slideSize.height } })
		} catch {
			slide.background = { color: model.background.fallback || 'FFFFFF' }
		}
	}

	for (const item of sortedItems(model.items)) {
		const p = item.position
		if (!p || p.w <= 0 || p.h < 0) continue
		// Render each item defensively: one bad graphic (e.g. an icon SVG that
		// fails to rasterize) must never abort the rest of the slide. Failures are
		// reported as warnings instead of silently swallowing the whole slide.
		try {
			if (item.type === 'image') {
				let src = item.src
				if (/^data:image\/svg\+xml/i.test(src)) src = await svgDataToPng(src, p.w * PX_PER_IN, p.h * PX_PER_IN)
				const opts: Record<string, unknown> = { ...imageRef(src), x: p.x, y: p.y, w: p.w, h: p.h }
				if (item.objectFit === 'cover') opts.sizing = { type: 'cover', w: p.w, h: p.h }
				if (item.objectFit === 'contain') opts.sizing = { type: 'contain', w: p.w, h: p.h }
				if (item.transparency != null) opts.transparency = item.transparency
				slide.addImage(opts)
			} else if (item.type === 'shape') {
				const rounded = item.radius && item.radius > 0
				const shape = rounded ? pptx.ShapeType.roundRect : pptx.ShapeType.rect
				const opts: Record<string, unknown> = { x: p.x, y: p.y, w: p.w, h: p.h }
				if (item.fill) {
					const fill: Record<string, unknown> = { color: item.fill }
					if (item.transparency != null) fill.transparency = item.transparency
					opts.fill = fill
				} else {
					opts.fill = { color: 'FFFFFF', transparency: 100 }
				}
				opts.line = item.line ?? { color: 'FFFFFF', transparency: 100 }
				slide.addShape(shape, opts)
			} else if (item.type === 'path') {
				// Editable freeform vector: SVG/icon paths vectorized into
				// custGeom instead of the PNG raster fallback above.
				emitPathItem(pptx, slide, item)
			} else if (item.type === 'line') {
				slide.addShape(pptx.ShapeType.line, {
					x: p.x,
					y: p.y,
					w: p.w,
					h: p.h || 0,
					line: { color: item.color || '888888', width: item.width || 1, dashType: item.dash || 'solid' },
				})
			} else if (item.type === 'table') {
				const opts: Record<string, unknown> = {
					x: p.x,
					y: p.y,
					w: p.w,
					h: p.h,
					border: item.border || { type: 'solid', color: 'FFFFFF', transparency: 100, pt: 0.25 },
					autoPage: false,
					valign: 'top',
				}
				if (item.colW && item.colW.length) opts.colW = item.colW
				if (item.rowH && item.rowH.length) opts.rowH = item.rowH
				slide.addTable(item.rows, opts)
			} else if (item.type === 'list' || item.type === 'text') {
				slide.addText(item.text, textOptions(item, p, slideSize))
			}
		} catch (err) {
			const what = item.type === 'image' ? (/^data:image\/svg/i.test(item.src || '') ? 'icon/SVG graphic' : 'image') : item.type
			issues.push('SLIDE EXPORT ISSUE: a ' + what + ' could not be placed (' + (err instanceof Error && err.message ? err.message : String(err)) + ').')
		}
	}
	if (model.notes && typeof slide.addNotes === 'function') slide.addNotes(model.notes)
	return issues
}

function textOptions(item: TextItem, p: Rect, slideSize: SlideSize): Record<string, unknown> {
	const s = item.style || {}
	const maxW = slideSize && Number.isFinite(slideSize.width) ? Math.max(0.05, slideSize.width - p.x) : null
	const maxH = slideSize && Number.isFinite(slideSize.height) ? Math.max(0.05, slideSize.height - p.y) : null
	const isHeading = /^h[1-6]$/i.test(item.tag || '')
	const constrain = !!s.constrainTextBox
	const textW = constrain ? p.w : p.w * (item.noWrap ? 1.65 : isHeading ? 1.0 : 1.14)
	const textH = constrain ? p.h : p.h * (item.noWrap ? 1.22 : 1.12)
	const boxH = constrain ? textH : textH + 0.08
	const opts: Record<string, unknown> = {
		x: p.x,
		y: p.y,
		w: Math.max(0.05, maxW ? Math.min(maxW, textW) : textW),
		h: Math.max(0.05, maxH ? Math.min(maxH, boxH) : boxH),
		fontFace: s.fontFace || DEFAULT_FONT.latin,
		fontSize: s.fontSize || 12,
		color: s.color || '000000',
		bold: !!s.bold,
		italic: !!s.italic,
		underline: !!s.underline,
		strike: !!s.strike,
		align: s.align || 'left',
		valign: s.valign || 'top',
		margin: s.margin || [0, 0, 0, 0],
		fit: 'none',
	}
	if (s.lineSpacing) opts.lineSpacing = s.lineSpacing
	if (s.transparency != null) opts.transparency = s.transparency
	if (item.noWrap) opts.wrap = false
	return opts
}
