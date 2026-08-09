/**
 * Emit layer — IR `path` item → ts-pptx `custGeom` shape.
 *
 * Pure and isomorphic (no DOM): this is the editable-vector counterpart of the
 * SVG→PNG raster fallback in `./slide.ts`. The IR's {@link FreeformPoint}
 * mirrors ts-pptx's freeform point DSL one-to-one, so points pass through
 * unchanged; coordinates are already in the shape box's own inch space (0..w,
 * 0..h), which is exactly what ts-pptx's `custGeom` emitter expects.
 *
 * Fill / stroke follow ts-pptx's defaults: omitting `fill` yields `<a:noFill/>`
 * (a stroke-only icon), and omitting `line` leaves the shape unstroked.
 */

import type { PathItem } from './model'
import type { PptxSlide, PptxWriter } from './pptx-types'

/** Build the ts-pptx `addShape('custGeom', …)` options for a path item. */
export function pathShapeOptions(item: PathItem): Record<string, unknown> {
	const p = item.position
	const opts: Record<string, unknown> = {
		x: p.x,
		y: p.y,
		w: p.w,
		h: p.h,
		points: item.points,
	}
	if (item.fill && item.fill !== 'none') opts.fill = { color: item.fill }
	if (item.line && item.line.color && item.line.color !== 'none') {
		opts.line = {
			color: item.line.color,
			width: item.line.width || 1,
			dashType: item.line.dash || 'solid',
		}
	}
	return opts
}

/**
 * Add a path item to a slide as an editable `custGeom` shape. `pptx` is the
 * writer instance (for `ShapeType.custGeom`); `slide` is the target slide.
 */
export function emitPathItem(pptx: PptxWriter, slide: Pick<PptxSlide, 'addShape'>, item: PathItem): void {
	slide.addShape(pptx.ShapeType.custGeom, pathShapeOptions(item))
}
