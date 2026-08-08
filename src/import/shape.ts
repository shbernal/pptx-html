/**
 * Read model → IR nodes: one mapper per shape kind, dispatched on the read
 * model's own guards.
 *
 * Two rules govern every mapper here.
 *
 * **No mapper returns an approximation.** A construct this IR does not model
 * becomes an {@link OpaqueNode} — a labelled inert box that says what it stands
 * for — and never a shape drawn to look roughly like it. An approximation passes
 * a glance and has no way back, which is the one output the charter rules out.
 *
 * **Every box is slide-absolute.** A group child's `a:xfrm` is stated in its
 * group's child space, so the read model's `absoluteFrame` composes the enclosing
 * chain and this file never does the arithmetic itself. That removes the trap
 * where a nested group's contents are all displaced by one scale factor — the
 * kind of bug a diff catches only if the corpus happens to contain a nested
 * group.
 */

import {
	type AnyShape,
	type Connector,
	type GraphicFrame,
	type GroupShape,
	isAutoShape,
	isConnector,
	isGraphicFrame,
	isGroupShape,
	isPicture,
	type Picture,
	type Shape,
	type Table,
	type TableCell as ReadTableCell,
} from '@shbernal/ts-pptx/read'
import {
	cellNodeId,
	type Connection,
	type EdgeRect,
	importedNodeId,
	type NodeId,
	type OpaqueNode,
	type Placement,
	type RenderNode,
	type TableCell,
	type TableColumn,
	type TableRow,
	type TextBody,
} from '../ir/render'
import { forShape, type ImportScope, note } from './context'
import { geometryOf } from './geometry'
import { cellBorderOf, fillOf, strokeOf } from './paint'
import { textBodyOf } from './text'

/** The identity every node on this slide is addressed by. */
function idOf(shape: Shape, scope: ImportScope): NodeId {
	// `p:cNvPr/@id` is required by the schema and unique within a shape tree, so a
	// missing one means a malformed package rather than a case to design for; `0`
	// keeps the id derivable instead of throwing on a deck we can still mostly draw.
	return importedNodeId(scope.slideNumber, shape.id ?? 0)
}

function placementOf(shape: Shape, scope: ImportScope): Placement | null {
	const absolute = shape.absoluteFrame
	const resolved = shape.resolvedFrame
	const geometrySource = resolved?.source ?? 'own'

	if (absolute !== null) {
		return {
			box: { x: absolute.left, y: absolute.top, w: absolute.width, h: absolute.height },
			rotation: absolute.rotation,
			flipH: absolute.flipH,
			flipV: absolute.flipV,
			geometrySource,
		}
	}
	if (resolved !== null) {
		// No composed frame but an inherited one: a placeholder taking its geometry
		// from the layout, which is exactly the case `geometrySource` exists to name.
		return {
			box: { x: resolved.left, y: resolved.top, w: resolved.width, h: resolved.height },
			rotation: shape.rotation ?? 0,
			flipH: shape.flipH,
			flipV: shape.flipV,
			geometrySource,
		}
	}

	note(
		scope,
		'shape.frameInherited',
		'dropped',
		'unread',
		'nothing in the slide, layout or master chain gives this shape a resolvable box, so it has no placement to draw at'
	)
	return null
}

/** The members every node shares, gathered once so no mapper can forget one. */
function baseOf(shape: Shape, scope: ImportScope): {
	id: NodeId
	name: string
	placement: Placement | null
	hidden?: boolean
	alt?: string
	placeholder?: { type: string | null; idx: string }
} {
	const placeholder = shape.placeholder
	const alt = shape.description
	return {
		id: idOf(shape, scope),
		name: shape.name,
		placement: placementOf(shape, scope),
		...(shape.hidden ? { hidden: true } : {}),
		...(alt === null || alt === '' ? {} : { alt }),
		// The link, not its resolution: a placeholder flattened into a plain shape
		// passes a pixel diff and loses the binding that governs theme and colour-map
		// resolution. Recording it is what keeps re-emitting against a layout possible.
		...(placeholder === null ? {} : { placeholder: { type: placeholder.type, idx: placeholder.idx } }),
	}
}

function opaque(shape: Shape, scope: ImportScope, standsFor: string): OpaqueNode {
	return { ...baseOf(shape, scope), kind: 'opaque', render: 'placeholder', standsFor }
}

function textOf(shape: Shape, scope: ImportScope): TextBody | null {
	if (!shape.hasTextFrame) return null
	const frame = shape.textFrame
	return frame === null ? null : textBodyOf(frame, scope)
}

export function nodeOf(shape: AnyShape, parent: ImportScope): RenderNode {
	const scope = forShape(parent, shape.name)
	if (isGroupShape(shape)) return groupNodeOf(shape, scope)
	if (isPicture(shape)) return pictureNodeOf(shape, scope)
	if (isConnector(shape)) return connectorNodeOf(shape, scope)
	if (isGraphicFrame(shape)) return graphicFrameNodeOf(shape, scope)
	if (isAutoShape(shape)) {
		return {
			...baseOf(shape, scope),
			kind: 'shape',
			render: 'drawn',
			geometry: geometryOf(shape, scope),
			fill: fillOf(shape, scope),
			stroke: strokeOf(shape, scope),
			text: textOf(shape, scope),
		}
	}
	// Unreachable through `AnyShape` today. Left as an opaque node rather than a
	// throw so a future read-model kind degrades to a labelled box instead of
	// taking the whole import down.
	return opaque(shape, scope, 'shape')
}

function groupNodeOf(group: GroupShape, scope: ImportScope): RenderNode {
	return {
		...baseOf(group, scope),
		kind: 'group',
		render: 'drawn',
		children: group.shapes.map((child) => nodeOf(child, scope)),
	}
}

function pictureNodeOf(picture: Picture, scope: ImportScope): RenderNode {
	const asset = scope.assets.refFor(picture.imagePartName)
	if (asset === null) {
		note(
			scope,
			'image.data',
			'dropped',
			'unread',
			'this picture resolves to no media part in the package, so there are no bytes to draw'
		)
		return opaque(picture, scope, 'picture')
	}
	const crop = picture.crop
	return {
		...baseOf(picture, scope),
		kind: 'picture',
		render: 'drawn',
		asset,
		...(crop === null ? {} : { crop: crop as EdgeRect }),
		geometry: geometryOf(picture, scope),
		stroke: strokeOf(picture, scope),
	}
}

function connectorNodeOf(connector: Connector, scope: ImportScope): RenderNode {
	const start = connectionOf(connector.startConnection, scope)
	const end = connectionOf(connector.endConnection, scope)
	return {
		...baseOf(connector, scope),
		kind: 'connector',
		render: 'drawn',
		geometry: geometryOf(connector, scope),
		stroke: strokeOf(connector, scope),
		...(start === null ? {} : { start }),
		...(end === null ? {} : { end }),
	}
}

/**
 * A bound connector end.
 *
 * `node` is `null` when the binding names a shape id nothing on the slide
 * carries. That is a dangling reference in the source, not a mapping failure, so
 * it is recorded rather than dropped — the site index still says where the line
 * was meant to attach.
 */
function connectionOf(
	site: { shapeId: number; siteIndex: number; boundShape: AnyShape | null } | null,
	scope: ImportScope
): Connection | null {
	if (site === null) return null
	if (site.boundShape === null) {
		note(
			scope,
			'connector.binding',
			'dropped',
			'unread',
			`this connector binds to shape id ${site.shapeId}, which is not on the slide`
		)
		return { node: null, site: site.siteIndex }
	}
	return { node: importedNodeId(scope.slideNumber, site.shapeId), site: site.siteIndex }
}

function graphicFrameNodeOf(frame: GraphicFrame, scope: ImportScope): RenderNode {
	if (frame.hasTable) {
		const table = frame.table
		if (table !== null) return tableNodeOf(frame, table, scope)
	}
	// Charts and their extended cousins are read in full and drawn by nothing:
	// there is no IR representation and no renderer, and inventing one would be an
	// approximation. They stand for themselves on the page.
	if (frame.hasChartEx) return opaque(frame, scope, 'chartEx')
	if (frame.hasChart) return opaque(frame, scope, 'chart')
	note(
		scope,
		'graphicFrame.unknown',
		'dropped',
		'unread',
		'this graphic frame holds neither a table nor a chart — SmartArt, an embedded object or a 3-D model — and the read model exposes no way in'
	)
	return opaque(frame, scope, 'graphicFrame')
}

function tableNodeOf(frame: GraphicFrame, table: Table, scope: ImportScope): RenderNode {
	const base = baseOf(frame, scope)
	const columns: TableColumn[] = table.columnWidths.map((widthEmu) => ({ widthEmu }))
	const rows: TableRow[] = table.rows.map((row, rowIndex) => ({
		heightEmu: row.heightEmu,
		cells: row.cells.map((cell, columnIndex) => cellOf(cell, base.id, rowIndex, columnIndex, scope)),
	}))
	return { ...base, kind: 'table', render: 'drawn', columns, rows }
}

const CELL_ANCHOR: Record<string, TableCell['anchor']> = { t: 'top', ctr: 'middle', b: 'bottom' }

function cellOf(
	cell: ReadTableCell,
	tableId: NodeId,
	rowIndex: number,
	columnIndex: number,
	scope: ImportScope
): TableCell {
	const frame = cell.textFrame
	const borders = cell.borders
	const margins = cell.marginsEmu
	const anchor = cell.anchor
	const spans = { columns: cell.gridSpan, rows: cell.rowSpan }
	return {
		id: cellNodeId(tableId, rowIndex, columnIndex),
		text: frame === null ? null : textBodyOf(frame, scope),
		fill: fillOf(cell, scope),
		borders: {
			left: cellBorderOf(borders?.left ?? null, scope),
			right: cellBorderOf(borders?.right ?? null, scope),
			top: cellBorderOf(borders?.top ?? null, scope),
			bottom: cellBorderOf(borders?.bottom ?? null, scope),
		},
		span: spans.columns > 1 || spans.rows > 1 ? spans : null,
		covered: cell.isMergeContinuation,
		// A cell with no `a:tcPr` margins at all and one that states only some
		// collapse to the same four nullable sides — `null` per side means the same
		// thing in both, so there is nothing lost in flattening the outer `null`.
		marginsEmu: {
			left: margins?.left ?? null,
			right: margins?.right ?? null,
			top: margins?.top ?? null,
			bottom: margins?.bottom ?? null,
		},
		anchor: (anchor !== null ? CELL_ANCHOR[anchor] : undefined) ?? 'top',
	}
}
