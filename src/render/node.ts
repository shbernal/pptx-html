/**
 * {@link RenderNode} → SVG, in slide space.
 *
 * Each node becomes one `<g>` carrying its {@link NodeId} and its placement
 * transform, and everything inside that group is drawn in the node's local
 * `0..w` × `0..h` EMU space. That split is what lets rotation and flips be a
 * single transform on the group instead of arithmetic threaded through every
 * path command and every text box.
 *
 * The group's `data-d2p-node` is a **link**, not state: it lets part 06 (and a
 * human with dev tools) find the island entry a painted element came from. The
 * model itself is never written into attributes — see `island.ts` for why.
 */

import type { Box, EdgeRect, Placement, RenderNode, TableNode } from '../ir/render'
import { pathOf } from './geometry'
import { type Defs, escapeAttr, fillPaint, strokePaint } from './paint'
import { renderTextBody } from './text'

export interface NodeContext {
	defs: Defs
	/** Nodes that could not be placed, for the caller to report. */
	skipped: string[]
}

/**
 * The placement transform.
 *
 * Order matters and is fixed by what each step means: translate the local space
 * to the box, then rotate about the box's own centre (`a:xfrm/@rot` is stated
 * about the shape's centre, not the slide's), then flip about that same centre.
 * A flip is `scale(-1)` sandwiched between two translates because SVG scales
 * about the origin and OOXML flips about the shape.
 */
function transformOf(placement: Placement): string {
	const { box, rotation, flipH, flipV } = placement
	const parts = [`translate(${box.x} ${box.y})`]
	if (rotation !== 0) parts.push(`rotate(${rotation} ${box.w / 2} ${box.h / 2})`)
	if (flipH || flipV) {
		parts.push(
			`translate(${flipH ? box.w : 0} ${flipV ? box.h : 0})`,
			`scale(${flipH ? -1 : 1} ${flipV ? -1 : 1})`
		)
	}
	return parts.join(' ')
}

/** A `<foreignObject>` covering the node's box. */
function textFrame(box: Box, html: string): string {
	return `<foreignObject x="0" y="0" width="${box.w}" height="${box.h}">${html}</foreignObject>`
}

/**
 * The labelled box a `render: 'placeholder'` node becomes.
 *
 * Deliberately not a lookalike. A drawn-on chart that round-trips as the real
 * chart is precisely the silent loss this project exists to remove, so the
 * placeholder is inert, obviously a placeholder, and names the construct it
 * stands in for.
 */
function placeholderBox(box: Box, standsFor: string): string {
	const label = escapeAttr(standsFor)
	return (
		`<rect x="0" y="0" width="${box.w}" height="${box.h}" fill="#f4f5f8" stroke="#9aa1b1" ` +
		`stroke-width="12700" stroke-dasharray="76200 38100"/>` +
		`<foreignObject x="0" y="0" width="${box.w}" height="${box.h}">` +
		`<div xmlns="http://www.w3.org/1999/xhtml" style="height:100%;display:flex;align-items:center;` +
		`justify-content:center;font:171450px sans-serif;color:#5b6272;text-align:center">` +
		`${label} — carried, not editable</div></foreignObject>`
	)
}

/**
 * A picture's `<image>`, with `a:srcRect` applied by oversizing and offsetting.
 *
 * A crop states how much of each edge to *discard*, so the surviving fraction is
 * `1 - left - right`; drawing the whole image that much larger and shifting it
 * left by the discarded part puts the intended window inside the box. Negative
 * insets bleed outward and fall out of the same arithmetic, which is why this is
 * computed rather than clamped.
 */
function croppedImage(asset: string, box: Box, crop: EdgeRect | undefined): string {
	const attrs = `data-d2p-asset="${escapeAttr(asset)}" preserveAspectRatio="none"`
	if (crop === undefined) return `<image ${attrs} x="0" y="0" width="${box.w}" height="${box.h}"/>`

	const visibleW = 1 - crop.left - crop.right
	const visibleH = 1 - crop.top - crop.bottom
	// A degenerate crop window would divide by zero and paint nothing useful;
	// falling back to the uncropped image keeps the picture recognisable.
	if (visibleW <= 0 || visibleH <= 0) return `<image ${attrs} x="0" y="0" width="${box.w}" height="${box.h}"/>`

	const width = box.w / visibleW
	const height = box.h / visibleH
	return `<image ${attrs} x="${-crop.left * width}" y="${-crop.top * height}" width="${width}" height="${height}"/>`
}

function renderTable(node: TableNode, box: Box, context: NodeContext): string {
	// A column with no stated width, or a row with no stated height, shares out
	// what the frame has left. That is the same rule the writer applies on the way
	// out, so a table drawn here and a table re-emitted agree about the geometry.
	const statedW = node.columns.reduce((sum, column) => sum + (column.widthEmu ?? 0), 0)
	const autoColumns = node.columns.filter((column) => column.widthEmu === null).length
	const columnFill = autoColumns === 0 ? 0 : Math.max(0, box.w - statedW) / autoColumns

	const statedH = node.rows.reduce((sum, row) => sum + (row.heightEmu ?? 0), 0)
	const autoRows = node.rows.filter((row) => row.heightEmu === null).length
	const rowFill = autoRows === 0 ? Math.max(0, box.h - statedH) / Math.max(1, node.rows.length) : Math.max(0, box.h - statedH) / autoRows

	const parts: string[] = []
	let y = 0
	for (const row of node.rows) {
		const h = row.heightEmu ?? rowFill
		let x = 0
		row.cells.forEach((cell, columnIndex) => {
			const w = node.columns[columnIndex]?.widthEmu ?? columnFill
			// A covered cell is the far side of a neighbour's span: it has no box of
			// its own, and painting one would draw a border through the merge.
			if (cell.covered) {
				x += w
				return
			}
			const span = cell.span
			const cellW = span === null ? w : node.columns.slice(columnIndex, columnIndex + span.columns).reduce((sum, column) => sum + (column.widthEmu ?? columnFill), 0)
			const cellH = span === null ? h : h * span.rows

			const fill = fillPaint(cell.fill, context.defs)
			parts.push(
				`<g data-d2p-node="${escapeAttr(cell.id)}" transform="translate(${x} ${y})"` +
					`${fill.approx === undefined ? '' : ` data-d2p-approx="${fill.approx}"`}>` +
					`<rect x="0" y="0" width="${cellW}" height="${cellH}" ${fill.attrs}/>` +
					edges(cell.borders, cellW, cellH, context) +
					(cell.text === null
						? ''
						: textFrame({ x: 0, y: 0, w: cellW, h: cellH }, renderTextBody(cell.text, cell.id))) +
					`</g>`
			)
			x += w
		})
		y += h
	}
	return parts.join('')
}

/** A cell's four edges, drawn as separate lines so each keeps its own stroke. */
function edges(borders: TableNode['rows'][number]['cells'][number]['borders'], w: number, h: number, context: NodeContext): string {
	const sides: [keyof typeof borders, number, number, number, number][] = [
		['top', 0, 0, w, 0],
		['right', w, 0, w, h],
		['bottom', 0, h, w, h],
		['left', 0, 0, 0, h],
	]
	return sides
		.map(([side, x1, y1, x2, y2]) => {
			const paint = strokePaint(borders[side], context.defs)
			if (paint.attrs === 'stroke="none"') return ''
			return `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" ${paint.attrs}/>`
		})
		.join('')
}

/** One node, and its children if it has any. */
export function renderNode(node: RenderNode, context: NodeContext): string {
	if (node.placement === null) {
		// Honest rather than placed at the origin: nothing in the slide → layout →
		// master chain gave this node a box, and inventing one would put a shape
		// somewhere the deck never said it was.
		context.skipped.push(node.id)
		return ''
	}

	const box = node.placement.box
	const attrs = [`data-d2p-node="${escapeAttr(node.id)}"`, `transform="${transformOf(node.placement)}"`]
	if (node.hidden === true) attrs.push('style="display:none"')
	if (node.alt !== undefined) attrs.push(`aria-label="${escapeAttr(node.alt)}"`)

	const inner = node.render === 'placeholder' ? placeholderBox(box, node.standsFor ?? node.kind) : bodyOf(node, box, context)

	return `<g ${attrs.join(' ')}>${inner}</g>`
}

function bodyOf(node: RenderNode, box: Box, context: NodeContext): string {
	switch (node.kind) {
		case 'group':
			// A group paints nothing of its own. Its children's placements are already
			// slide-absolute, so they are rendered against the slide and not against
			// the group — which is why this reaches past its own transform.
			return node.children.map((child) => renderNode(child, context)).join('')

		case 'shape': {
			const path = pathOf(node.geometry, box.w, box.h)
			const fill = path.unfilled === true ? { attrs: 'fill="none"' } : fillPaint(node.fill, context.defs)
			const stroke = path.unstroked === true ? { attrs: 'stroke="none"' } : strokePaint(node.stroke, context.defs)
			const approx = [path.fallback === undefined ? '' : `geometry:${path.fallback}`, fill.approx ?? '', stroke.approx ?? '']
				.filter((entry) => entry !== '')
				.join(' ')
			return (
				`<path d="${path.d}" ${fill.attrs} ${stroke.attrs}` +
				`${approx === '' ? '' : ` data-d2p-approx="${escapeAttr(approx)}"`}/>` +
				(node.text === null ? '' : textFrame(box, renderTextBody(node.text, node.id)))
			)
		}

		case 'picture': {
			const path = pathOf(node.geometry, box.w, box.h)
			const stroke = strokePaint(node.stroke, context.defs)
			// The geometry clips the image rather than outlining it: a picture cropped
			// to a shape is the common case, and a plain `rect` clip is a no-op.
			const clip = context.defs.add((id) => `<clipPath id="${id}"><path d="${path.d}"/></clipPath>`)
			return (
				`<g clip-path="${clip}">${croppedImage(node.asset.$asset, box, node.crop)}</g>` +
				(stroke.attrs === 'stroke="none"' ? '' : `<path d="${path.d}" fill="none" ${stroke.attrs}/>`)
			)
		}

		case 'connector': {
			const path = pathOf(node.geometry, box.w, box.h)
			const stroke = strokePaint(node.stroke, context.defs)
			return `<path d="${path.d}" fill="none" ${stroke.attrs}${stroke.approx === undefined ? '' : ` data-d2p-approx="${stroke.approx}"`}/>`
		}

		case 'table':
			return renderTable(node, box, context)

		case 'opaque':
			// Unreachable in practice — an opaque node is always `render:
			// 'placeholder'` and `renderNode` has already branched — but the union is
			// closed and the compiler is right to want the arm.
			return placeholderBox(box, node.standsFor)
	}
}
