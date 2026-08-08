/**
 * A `RenderIr` exercising every node kind, every fill kind, both geometry kinds
 * and both bullet kinds — hand-written, because the import layer that would
 * produce one does not exist yet.
 *
 * It is deliberately built with **no `undefined` values**: an absent field is a
 * missing key. The JSON island depends on that being true, and a fixture that
 * quietly wrote `undefined` would let the identity test pass while the real
 * model failed.
 */

import type {
	Color,
	ConnectorNode,
	Fill,
	GroupNode,
	OpaqueNode,
	PictureNode,
	RenderIr,
	ShapeNode,
	Stroke,
	TableNode,
	TextBody,
} from '../../src/ir/render'
import { cellNodeId, EMU_PER_INCH, IR_VERSION, importedNodeId } from '../../src/ir/render'

const inch = (n: number): number => Math.round(n * EMU_PER_INCH)

const ACCENT: Color = {
	kind: 'scheme',
	slot: 'accent1',
	transforms: [{ name: 'lumMod', value: '75000' }],
	effectiveHex: '2E5A8A',
}

const INK: Color = { kind: 'srgb', hex: '202020' }

const HAIRLINE: Stroke = {
	widthPt: 1,
	color: INK,
	dash: 'solid',
	cap: 'flat',
	join: 'miter',
}

const SOLID: Fill = { kind: 'solid', color: ACCENT }

function body(text: string): TextBody {
	return {
		paragraphs: [
			{
				props: { align: 'left', level: 0, bullet: { kind: 'none' } },
				runs: [{ text, props: { bold: true, sizePt: 18, color: INK } }],
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
	placement: { box: { x: inch(0.5), y: inch(0.4), w: inch(9), h: inch(1.2) }, rotation: 0, flipH: false, flipV: false },
	render: 'drawn',
	placeholder: { type: 'title', idx: '0' },
	geometry: { kind: 'preset', preset: 'rect', adjustValues: {} },
	fill: { kind: 'none' },
	stroke: null,
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
					{ text: 'Quarterly ', props: { sizePt: 40, color: ACCENT } },
					{ text: 'review', props: { sizePt: 40, italic: true, color: ACCENT, underline: 'single' } },
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
	placement: { box: { x: inch(0.5), y: inch(1.8), w: inch(5), h: inch(3) }, rotation: 0, flipH: false, flipV: false },
	render: 'drawn',
	geometry: { kind: 'preset', preset: 'roundRect', adjustValues: { adj: 16667 } },
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
	stroke: { ...HAIRLINE, dash: 'dash', head: { type: 'none', width: 'med', length: 'med' } },
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
				runs: [{ text: 'Revenue up', props: {} }],
			},
			{
				props: { align: 'left', level: 1, bullet: { kind: 'number', scheme: 'arabicPeriod', startAt: 1 } },
				runs: [{ text: 'North', props: { bold: true } }],
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
	placement: { box: { x: inch(6), y: inch(2), w: inch(2), h: inch(2) }, rotation: 30, flipH: true, flipV: false },
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
	placement: { box: { x: inch(8), y: inch(0.4), w: inch(1), h: inch(1) }, rotation: 0, flipH: false, flipV: false },
	render: 'drawn',
	alt: 'Company logo',
	asset: { $asset: 'image1.png' },
	crop: { left: 0.1, top: 0, right: 0.1, bottom: 0 },
	geometry: { kind: 'preset', preset: 'rect', adjustValues: {} },
	stroke: null,
}

const ARROW: ConnectorNode = {
	kind: 'connector',
	id: importedNodeId(1, 6),
	name: 'Straight Arrow Connector 5',
	placement: { box: { x: inch(5.6), y: inch(3), w: inch(0.4), h: inch(0) }, rotation: 0, flipH: false, flipV: false },
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
	placement: { box: { x: inch(0.5), y: inch(5), w: inch(3), h: inch(1) }, rotation: 0, flipH: false, flipV: true },
	render: 'drawn',
	childOffset: { x: 0, y: 0 },
	childExtent: { w: inch(3), h: inch(1) },
	children: [
		{
			kind: 'shape',
			id: importedNodeId(1, 8),
			name: 'Chip 7',
			placement: { box: { x: 0, y: 0, w: inch(1.4), h: inch(1) }, rotation: 0, flipH: false, flipV: false },
			render: 'drawn',
			geometry: { kind: 'preset', preset: 'ellipse', adjustValues: {} },
			fill: SOLID,
			stroke: null,
			text: body('Nested'),
		},
	],
}

const TABLE_ID = importedNodeId(2, 2)

const GRID: TableNode = {
	kind: 'table',
	id: TABLE_ID,
	name: 'Table 1',
	placement: { box: { x: inch(0.5), y: inch(1), w: inch(6), h: inch(1.2) }, rotation: 0, flipH: false, flipV: false },
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
					marginsPt: { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 },
					anchor: 'middle',
				},
				{
					id: cellNodeId(TABLE_ID, 0, 1),
					text: null,
					fill: { kind: 'none' },
					borders: { left: null, right: null, top: null, bottom: null },
					span: null,
					covered: true,
					marginsPt: { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 },
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
	placement: { box: { x: inch(0.5), y: inch(2.6), w: inch(6), h: inch(3) }, rotation: 0, flipH: false, flipV: false },
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
			background: { source: 'slide', fill: { kind: 'picture', asset: { $asset: 'image1.png' } } },
			nodes: [GRID, CHART],
			notes: null,
			residual: { kind: 'slide', xml: '<p:sld><p:cSld/></p:sld>', assets: [{ $asset: 'image1.png' }] },
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
