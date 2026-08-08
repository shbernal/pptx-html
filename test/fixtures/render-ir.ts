/**
 * A `RenderIr` exercising every node kind, every fill and stroke state, both
 * geometry kinds and both bullet kinds.
 *
 * Still hand-written now that `src/import` exists, and deliberately so: an
 * imported fixture would test the model against whatever the importer happens to
 * produce, so a field the importer never populates would look covered. This one
 * is the model's own statement of what it can hold, and `test/oracle/` is where
 * the importer is held to producing it.
 *
 * It is built with **no `undefined` values**: an absent field is a missing key.
 * The JSON island depends on that being true, and a fixture that quietly wrote
 * `undefined` would let the identity test pass while the real model failed.
 */

import type {
	Color,
	ConnectorNode,
	Fill,
	GroupNode,
	OpaqueNode,
	PictureNode,
	Placement,
	RenderIr,
	ShapeNode,
	Stroke,
	TableNode,
	TextBody,
} from '../../src/ir/render'
import { cellNodeId, EMU_PER_INCH, IR_VERSION, importedNodeId } from '../../src/ir/render'

const inch = (n: number): number => Math.round(n * EMU_PER_INCH)

/** The drawn arm of {@link Stroke}, so spreading one keeps its discriminant. */
type Line = Extract<Stroke, { kind: 'line' }>

const ACCENT: Color = {
	kind: 'scheme',
	slot: 'accent1',
	transforms: [{ name: 'lumMod', value: '75000' }],
	effectiveHex: '2E5A8A',
}

const INK: Color = { kind: 'srgb', hex: '202020' }

const HAIRLINE: Line = { kind: 'line', widthPt: 1, color: INK, dash: 'solid' }

const INHERITED: Stroke = { kind: 'inherit' }

const SOLID: Fill = { kind: 'solid', color: ACCENT }

/** Every node's box is slide-absolute; only the placeholder tier varies. */
function at(x: number, y: number, w: number, h: number, rest: Partial<Placement> = {}): Placement {
	return {
		box: { x: inch(x), y: inch(y), w: inch(w), h: inch(h) },
		rotation: 0,
		flipH: false,
		flipV: false,
		geometrySource: 'own',
		...rest,
	}
}

function body(text: string): TextBody {
	return {
		paragraphs: [
			{
				props: { align: 'left', level: 0, bullet: { kind: 'none' } },
				runs: [{ text, props: { bold: true, sizePt: 18, color: INK }, resolved: { bold: true, sizePt: 18 } }],
			},
		],
		autofit: 'none',
		anchor: 'top',
		wrap: true,
		insetsPt: { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 },
	}
}

const TITLE: ShapeNode = {
	kind: 'shape',
	id: importedNodeId(1, 2),
	name: 'Title 1',
	// The one node whose geometry comes from the layout rather than the slide —
	// the case `geometrySource` exists to keep visible.
	placement: at(0.5, 0.4, 9, 1.2, { geometrySource: 'layout' }),
	render: 'drawn',
	placeholder: { type: 'title', idx: '0' },
	geometry: { kind: 'preset', preset: 'rect', adjustValues: {} },
	fill: { kind: 'none' },
	stroke: INHERITED,
	text: {
		paragraphs: [
			{
				props: {
					align: 'center',
					level: 0,
					bullet: { kind: 'none' },
					lineSpacing: { type: 'percent', percent: 150 },
					spaceAfterPt: 6,
				},
				runs: [
					{
						text: 'Quarterly ',
						props: { sizePt: 40, color: ACCENT },
						resolved: { sizePt: 40, color: { kind: 'srgb', hex: '2E5A8A' } },
					},
					{
						text: 'review',
						props: { sizePt: 40, italic: true, color: ACCENT, underline: 'single' },
						resolved: { sizePt: 40, color: { kind: 'srgb', hex: '2E5A8A' } },
					},
				],
			},
		],
		autofit: 'shrink',
		anchor: 'middle',
		wrap: true,
		insetsPt: { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 },
	},
}

const BULLETS: ShapeNode = {
	kind: 'shape',
	id: importedNodeId(1, 3),
	name: 'Content Placeholder 2',
	placement: at(0.5, 1.8, 5, 3),
	render: 'drawn',
	geometry: { kind: 'preset', preset: 'roundRect', adjustValues: { adj: 'val 16667' } },
	fill: {
		kind: 'gradient',
		gradient: {
			kind: 'linear',
			angleDeg: 45,
			stops: [
				{ position: 0, color: ACCENT },
				{ position: 1, color: INK },
			],
		},
	},
	stroke: { ...HAIRLINE, dash: 'dash', head: { type: 'none' } },
	text: {
		paragraphs: [
			{
				props: {
					align: 'left',
					level: 0,
					bullet: { kind: 'character', char: '•', font: 'Arial' },
					marginLeftPt: 18,
					indentPt: -18,
				},
				// The inherited run: it states nothing, so everything a renderer needs
				// is in `resolved` and `props` is empty.
				runs: [{ text: 'Revenue up', props: {}, resolved: { sizePt: 18, fontFace: 'Calibri' } }],
			},
			{
				props: { align: 'left', level: 1, bullet: { kind: 'number', scheme: 'arabicPeriod', startAt: 1 } },
				runs: [{ text: 'North', props: { bold: true }, resolved: { bold: true, sizePt: 18 } }],
			},
		],
		autofit: 'resize',
		anchor: 'top',
		wrap: true,
		insetsPt: { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 },
		vertical: 'vert270',
	},
}

const FREEFORM: ShapeNode = {
	kind: 'shape',
	id: importedNodeId(1, 4),
	name: 'Freeform 3',
	placement: at(6, 2, 2, 2, { rotation: 30, flipH: true }),
	render: 'drawn',
	geometry: {
		kind: 'custom',
		paths: [
			{
				w: 100,
				h: 100,
				fill: 'norm',
				stroke: true,
				commands: [
					{ cmd: 'moveTo', x: 0, y: 0 },
					{ cmd: 'lnTo', x: 100, y: 0 },
					{ cmd: 'cubicBezTo', x1: 100, y1: 50, x2: 50, y2: 100, x: 0, y: 100 },
					{ cmd: 'close' },
				],
			},
		],
	},
	fill: { kind: 'pattern', preset: 'pct50', foreground: ACCENT, background: null },
	stroke: HAIRLINE,
	text: null,
}

const LOGO: PictureNode = {
	kind: 'picture',
	id: importedNodeId(1, 5),
	name: 'Picture 4',
	placement: at(8, 0.4, 1, 1),
	render: 'drawn',
	alt: 'Company logo',
	asset: { $asset: 'image1.png' },
	crop: { left: 0.1, top: 0, right: 0.1, bottom: 0 },
	geometry: { kind: 'preset', preset: 'rect', adjustValues: {} },
	stroke: { kind: 'none' },
}

const ARROW: ConnectorNode = {
	kind: 'connector',
	id: importedNodeId(1, 6),
	name: 'Straight Arrow Connector 5',
	placement: at(5.6, 3, 0.4, 0),
	render: 'drawn',
	geometry: { kind: 'preset', preset: 'straightConnector1', adjustValues: {} },
	stroke: { ...HAIRLINE, tail: { type: 'triangle', width: 'med', length: 'med' } },
	start: { node: importedNodeId(1, 3), site: 3 },
	end: { node: null, site: 1 },
}

const GROUP: GroupNode = {
	kind: 'group',
	id: importedNodeId(1, 7),
	name: 'Group 6',
	placement: at(0.5, 5, 3, 1, { flipV: true }),
	render: 'drawn',
	children: [
		{
			kind: 'shape',
			id: importedNodeId(1, 8),
			name: 'Chip 7',
			// Slide-absolute, inside the group's box — not child-space coordinates.
			placement: at(0.6, 5.1, 1.4, 0.8),
			render: 'drawn',
			geometry: { kind: 'preset', preset: 'ellipse', adjustValues: {} },
			fill: SOLID,
			stroke: INHERITED,
			text: body('Nested'),
		},
	],
}

const TABLE_ID = importedNodeId(2, 2)

const CELL_MARGINS = { left: 91440, right: 91440, top: 45720, bottom: 45720 }

const GRID: TableNode = {
	kind: 'table',
	id: TABLE_ID,
	name: 'Table 1',
	placement: at(0.5, 1, 6, 1.2),
	render: 'drawn',
	columns: [{ widthEmu: inch(3) }, { widthEmu: inch(3) }],
	rows: [
		{
			heightEmu: inch(0.6),
			cells: [
				{
					id: cellNodeId(TABLE_ID, 0, 0),
					text: body('Region'),
					fill: SOLID,
					borders: { left: HAIRLINE, right: HAIRLINE, top: HAIRLINE, bottom: HAIRLINE },
					span: { columns: 2, rows: 1 },
					covered: false,
					marginsEmu: CELL_MARGINS,
					anchor: 'middle',
				},
				{
					id: cellNodeId(TABLE_ID, 0, 1),
					text: null,
					fill: { kind: 'inherit' },
					borders: { left: INHERITED, right: INHERITED, top: INHERITED, bottom: INHERITED },
					span: null,
					covered: true,
					marginsEmu: { left: null, right: null, top: null, bottom: null },
					anchor: 'top',
				},
			],
		},
	],
}

const CHART: OpaqueNode = {
	kind: 'opaque',
	id: importedNodeId(2, 3),
	name: 'Chart 2',
	placement: at(0.5, 2.6, 6, 3),
	render: 'placeholder',
	standsFor: 'chart',
}

/**
 * Two slides: one `authored` and fully modeled, one `carried` whose chart the
 * model does not represent — so it keeps its source XML in the residual channel
 * and declares the loss.
 */
export const SAMPLE_IR: RenderIr = {
	irVersion: IR_VERSION,
	size: { w: inch(10), h: inch(5.625) },
	assets: [
		{
			name: 'image1.png',
			contentType: 'image/png',
			byteLength: 68,
			sha256: '5891b5b522d5df086d0ff0b110fbd9d21bb4fc7163af34d08286a2e846f6be03',
		},
	],
	slides: [
		{
			number: 1,
			source: 'authored',
			layout: { name: 'Title and Content', index: 1, nameIsUnique: true },
			hidden: false,
			background: { source: 'master', fill: { kind: 'solid', color: { kind: 'srgb', hex: 'FFFFFF' } } },
			nodes: [TITLE, BULLETS, FREEFORM, LOGO, ARROW, GROUP],
			notes: 'Open with the headline number.',
			residual: null,
			fidelity: [],
		},
		{
			number: 2,
			source: 'carried',
			layout: { name: 'Title Only', index: 2, nameIsUnique: true },
			hidden: true,
			background: { source: 'slide', fill: { kind: 'picture', asset: { $asset: 'image1.png' }, mode: 'stretch' } },
			nodes: [GRID, CHART],
			notes: null,
			residual: {
				kind: 'slide',
				xml: '<p:sld><p:cSld/></p:sld>',
				assets: [{ relId: 'rId2', asset: { $asset: 'image1.png' } }],
			},
			fidelity: [
				{
					slideNumber: 2,
					shapeName: 'Chart 2',
					construct: 'chart.workbook',
					disposition: 'approximated',
					cause: 'unsupported',
					detail: 'the render model has no chart node, so the slide is carried and the chart is drawn as a placeholder',
				},
			],
		},
	],
}
