/**
 * The probe manifest: writer options sent through the loop one at a time.
 *
 * Each entry is one option set to one value on an otherwise plain deck (see
 * `probe.ts`). The coverage loop grows this file one probe per commit; the
 * frontier it works from is `pnpm run oracle:score`, which lists every inventory
 * path that neither a probe nor `SKIPPED` accounts for.
 *
 * Rules a probe must keep, each of which a test enforces or the score exposes:
 *
 * - **Change one thing.** `author(pptx, false)` and `author(pptx, true)` differ by
 *   the option under test and nothing else, or the signature measures the wrong
 *   thing. Prefer the helpers below, which hold that by construction.
 * - **Use a non-default value.** A default writes nothing and the probe is inert.
 * - **Never edit a probe to make it pass.** A loss is recorded with `issue` (the
 *   upstream lane loses it) or `accepted` (with a reason), never by changing the
 *   value until it happens to survive.
 */

import type TsPptx from 'pptx-ts'
import type { SHAPE_NAME, ShapeProps, TableCellProps, TableProps, TextPropsOptions } from 'pptx-ts'
import type { Probe } from './probe'

const BOX = { x: 1, y: 1, w: 5, h: 1.5 }

/** `addText` on one box. `options` is applied only when the probe is on. */
function text(id: string, options: TextPropsOptions, covers = [id.split(':')[0] as string]): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addText('Probe text', { ...BOX, objectName: 'probe', ...(on ? options : {}) })
		},
	}
}

/** `addText` on one box with `base` on both sides, for options that only mean something on top of another. */
function textOn(
	id: string,
	base: TextPropsOptions,
	options: TextPropsOptions,
	covers = [id.split(':')[0] as string]
): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addText('Probe text', { ...BOX, objectName: 'probe', ...base, ...(on ? options : {}) })
		},
	}
}

/** `addText` with two runs; `options` go on run `at` only. */
function run(id: string, at: 0 | 1, options: TextPropsOptions, covers = [id.split(':')[0] as string]): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addText(
				['First run', 'Second run'].map((text, i) => ({ text, options: on && i === at ? options : {} })),
				{ ...BOX, objectName: 'probe' }
			)
		},
	}
}

/** A one-pixel PNG, for picture fills and picture bullets. */
const PNG =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

/** `addShape` with a visible fill, so geometry and line options have something to sit on. */
function shape(id: string, options: ShapeProps, covers = [id.split(':')[0] as string]): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			pptx
				.addSlide()
				.addShape('rect', { ...BOX, fill: { color: 'DDE3F0' }, objectName: 'probe', ...(on ? options : {}) })
		},
	}
}

/** `addShape` of preset `kind` with `base` on both sides, for options that only mean something on top of another. */
function shapeOn(
	id: string,
	kind: SHAPE_NAME,
	base: ShapeProps,
	options: ShapeProps,
	covers = [id.split(':')[0] as string]
): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addShape(kind, {
				...BOX,
				fill: { color: 'DDE3F0' },
				objectName: 'probe',
				...base,
				...(on ? options : {}),
			})
		},
	}
}

/** A closed triangle in the probe box, for custGeom options that need a path to sit on. */
const CUST = [{ x: 0, y: 1.5 }, { x: 2.5, y: 0 }, { x: 5, y: 1.5 }, { close: true as const }]

/** A 2×2 table; `options` go on the table call. */
function table(id: string, options: TableProps, covers = [id.split(':')[0] as string]): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addTable(
				[
					[{ text: 'a' }, { text: 'b' }],
					[{ text: 'c' }, { text: 'd' }],
				],
				{ ...BOX, objectName: 'probe', ...(on ? options : {}) }
			)
		},
	}
}

/** A 2×2 table; `options` go on the first cell only. */
function cell(id: string, options: TableCellProps, covers = [id.split(':')[0] as string]): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addTable(
				[
					[{ text: 'a', options: on ? options : {} }, { text: 'b' }],
					[{ text: 'c' }, { text: 'd' }],
				],
				{ ...BOX, objectName: 'probe' }
			)
		},
	}
}

export const PROBES: Probe[] = [
	text('addText.bold', { bold: true }),
	text('addText.italic', { italic: true }),
	text('addText.underline.style:dbl', { underline: { style: 'dbl' } }),
	text('addText.strike:dblStrike', { strike: 'dblStrike' }),
	{ ...text('addText.charSpacing', { charSpacing: 4 }), issue: 'https://github.com/shbernal/pptx-ts/issues/53' },
	text('addText.lineSpacingMultiple', { lineSpacingMultiple: 1.5 }),
	text('addText.valign:bottom', { valign: 'bottom' }),
	text('addText.margin', { margin: [10, 20, 10, 20] }),
	shape('addShape.line.dashType:dash', { line: { color: '250F6B', width: 2, dashType: 'dash' } }, [
		'addShape.line.dashType',
	]),
	shape('addShape.rotate', { rotate: 30 }),
	table('addTable.border.color', { border: { type: 'solid', color: '9B1B30', width: 2 } }, ['addTable.border.color']),
	cell('addTable.cell.fill.color', { fill: { color: 'F0C808' } }),
	text('addText.align:center', { align: 'center' }),
	text('addText.altText', { altText: 'A probe description' }),
	textOn('addText.angleRange', { shape: 'blockArc' }, { angleRange: [30, 200] }),
	textOn('addText.arcThicknessRatio', { shape: 'blockArc', angleRange: [30, 200] }, { arcThicknessRatio: 0.3 }),
	{ ...text('addText.baseline', { baseline: 30 }), issue: 'https://github.com/shbernal/pptx-ts/issues/55' },
	run('addText.breakLine', 0, { breakLine: true }),
	text('addText.bullet', { bullet: true }),
	textOn('addText.bullet.characterCode', { bullet: true }, { bullet: { characterCode: '25BA' } }),
	textOn('addText.bullet.color', { bullet: true }, { bullet: { color: '9B1B30' } }),
	textOn('addText.bullet.fontFace', { bullet: true }, { bullet: { fontFace: 'Wingdings' } }),
	{
		...textOn('addText.bullet.image.data', { bullet: true }, { bullet: { image: { data: PNG } } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/58',
	},
	textOn('addText.bullet.indent', { bullet: true }, { bullet: { indent: 10 } }),
	textOn(
		'addText.bullet.numberStartAt',
		{ bullet: { type: 'number' } },
		{ bullet: { type: 'number', numberStartAt: 5 } }
	),
	textOn(
		'addText.bullet.numberType',
		{ bullet: { type: 'number' } },
		{ bullet: { type: 'number', numberType: 'romanLcParenR' } }
	),
	textOn('addText.bullet.size', { bullet: true }, { bullet: { size: 80 } }),
	text('addText.bullet.type:number', { bullet: { type: 'number' } }),
	text('addText.caps:all', { caps: 'all' }),
	text('addText.color', { color: '9B1B30' }),
	{ ...text('addText.columns', { columns: 2 }), issue: 'https://github.com/shbernal/pptx-ts/issues/54' },
	{
		...textOn('addText.columnSpacing', { columns: 2 }, { columnSpacing: 0.5 }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/54',
	},
	text('addText.fill.color', { fill: { color: 'F0C808' } }, ['addText.fill', 'addText.fill.color']),
	{
		...textOn(
			'addText.fill.transparency',
			{ fill: { color: 'F0C808' } },
			{ fill: { color: 'F0C808', transparency: 40 } }
		),
		issue: 'https://github.com/shbernal/pptx-ts/issues/56',
	},
	{
		...text('addText.fill.type:inherit', { fill: { type: 'inherit' } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/56',
	},
	text(
		'addText.fill.gradient.linear',
		{
			fill: {
				gradient: {
					kind: 'linear',
					angle: 45,
					stops: [
						{ position: 0, color: 'F0C808' },
						{ position: 100, color: '250F6B', transparency: 30 },
					],
				},
			},
		},
		[
			'addText.fill.gradient.kind',
			'addText.fill.gradient.angle',
			'addText.fill.gradient.stops.color',
			'addText.fill.gradient.stops.position',
			'addText.fill.gradient.stops.transparency',
		]
	),
	{
		...textOn(
			'addText.fill.gradient.scaled',
			{
				fill: {
					gradient: {
						kind: 'linear',
						angle: 45,
						stops: [
							{ position: 0, color: 'F0C808' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			},
			{
				fill: {
					gradient: {
						kind: 'linear',
						angle: 45,
						scaled: true,
						stops: [
							{ position: 0, color: 'F0C808' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			}
		),
		issue: 'https://github.com/shbernal/pptx-ts/issues/56',
	},
	{
		...textOn(
			'addText.fill.gradient.rotateWithShape',
			{
				fill: {
					gradient: {
						kind: 'linear',
						angle: 45,
						stops: [
							{ position: 0, color: 'F0C808' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			},
			{
				fill: {
					gradient: {
						kind: 'linear',
						angle: 45,
						rotateWithShape: false,
						stops: [
							{ position: 0, color: 'F0C808' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			}
		),
		issue: 'https://github.com/shbernal/pptx-ts/issues/56',
	},
	{
		...text(
			'addText.fill.gradient.center',
			{
				fill: {
					gradient: {
						kind: 'radial',
						center: { x: 20, y: 70 },
						stops: [
							{ position: 0, color: 'F0C808' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			},
			['addText.fill.gradient.center.x', 'addText.fill.gradient.center.y']
		),
		issue: 'https://github.com/shbernal/pptx-ts/issues/56',
	},
	text('addText.fill.pattern', { fill: { pattern: { preset: 'dkDnDiag', fgColor: '250F6B', bgColor: 'F0C808' } } }, [
		'addText.fill.pattern.preset',
		'addText.fill.pattern.fgColor',
		'addText.fill.pattern.bgColor',
	]),
	text('addText.fill.image.data', { fill: { image: { data: PNG } } }),
	textOn(
		'addText.fill.image.crop',
		{ fill: { image: { data: PNG } } },
		{ fill: { image: { data: PNG, crop: { l: 10, t: 20, r: 30, b: 5 } } } },
		['addText.fill.image.crop.l', 'addText.fill.image.crop.t', 'addText.fill.image.crop.r', 'addText.fill.image.crop.b']
	),
	text('addText.fit:shrink', { fit: 'shrink' }),
	text('addText.fit:resize', { fit: 'resize' }),
	text('addText.fit.shrink', { fit: { type: 'shrink', fontScale: 80, lnSpcReduction: 10 } }, [
		'addText.fit.type',
		'addText.fit.fontScale',
		'addText.fit.lnSpcReduction',
	]),
	text('addText.flipH', { flipH: true }),
	text('addText.flipV', { flipV: true }),
	text('addText.fontFace', { fontFace: 'Georgia' }),
	{
		...textOn('addText.fontFaceEA', { fontFace: 'Georgia' }, { fontFaceEA: 'MS Mincho' }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/54',
	},
	text('addText.fontSize', { fontSize: 28 }),
	{
		...text('addText.glow', { glow: { size: 8, opacity: 0.5, color: 'FF00FF' } }, [
			'addText.glow.size',
			'addText.glow.opacity',
			'addText.glow.color',
		]),
		issue: 'https://github.com/shbernal/pptx-ts/issues/54',
	},
	text('addText.x', { x: 2 }),
	text('addText.y', { y: 2 }),
	text('addText.w', { w: 6 }),
	text('addText.h', { h: 2 }),
	text('addText.highlight', { highlight: 'FFFF00' }),
	{
		...text('addText.hyperlink.url', { hyperlink: { url: 'https://example.com/', tooltip: 'Example' } }, [
			'addText.hyperlink.url',
			'addText.hyperlink.tooltip',
		]),
		issue: 'https://github.com/shbernal/pptx-ts/issues/59',
	},
	{
		...text('addText.hyperlink.slide', { hyperlink: { slide: 1 } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/59',
	},
	{
		...text('addText.hyperlink.action:nextslide', { hyperlink: { action: 'nextslide' } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/59',
	},
	text('addText.indentLevel', { indentLevel: 2 }),
	text('addText.isTextBox', { isTextBox: true }),
	{ ...text('addText.lang', { lang: 'fr-FR' }), issue: 'https://github.com/shbernal/pptx-ts/issues/54' },
	text('addText.line', { line: { color: '250F6B', width: 2 } }, ['addText.line.color', 'addText.line.width']),
	textOn(
		'addText.line.dashType:dash',
		{ line: { color: '250F6B', width: 2 } },
		{ line: { color: '250F6B', width: 2, dashType: 'dash' } }
	),
	textOn(
		'addText.line.cap:round',
		{ line: { color: '250F6B', width: 2 } },
		{ line: { color: '250F6B', width: 2, cap: 'round' } }
	),
	textOn(
		'addText.line.beginArrowType',
		{ line: { color: '250F6B', width: 2 } },
		{ line: { color: '250F6B', width: 2, beginArrowType: 'triangle' } }
	),
	textOn(
		'addText.line.endArrowType',
		{ line: { color: '250F6B', width: 2 } },
		{ line: { color: '250F6B', width: 2, endArrowType: 'oval' } }
	),
	textOn(
		'addText.line.transparency',
		{ line: { color: '250F6B', width: 2 } },
		{ line: { color: '250F6B', width: 2, transparency: 40 } }
	),
	textOn(
		'addText.line.type:none',
		{ line: { color: '250F6B', width: 2 } },
		{ line: { color: '250F6B', width: 2, type: 'none' } }
	),
	{
		...text(
			'addText.line.gradient.linear',
			{
				line: {
					width: 3,
					gradient: {
						kind: 'linear',
						angle: 45,
						stops: [
							{ position: 0, color: 'F0C808' },
							{ position: 100, color: '250F6B', transparency: 30 },
						],
					},
				},
			},
			[
				'addText.line.gradient.kind',
				'addText.line.gradient.angle',
				'addText.line.gradient.stops.color',
				'addText.line.gradient.stops.position',
				'addText.line.gradient.stops.transparency',
			]
		),
		issue: 'https://github.com/shbernal/pptx-ts/issues/56',
	},
	{
		...textOn(
			'addText.line.gradient.scaled',
			{
				line: {
					width: 3,
					gradient: {
						kind: 'linear',
						angle: 45,
						stops: [
							{ position: 0, color: 'F0C808' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			},
			{
				line: {
					width: 3,
					gradient: {
						kind: 'linear',
						angle: 45,
						scaled: true,
						stops: [
							{ position: 0, color: 'F0C808' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			}
		),
		issue: 'https://github.com/shbernal/pptx-ts/issues/56',
	},
	{
		...textOn(
			'addText.line.gradient.rotateWithShape',
			{
				line: {
					width: 3,
					gradient: {
						kind: 'linear',
						angle: 45,
						stops: [
							{ position: 0, color: 'F0C808' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			},
			{
				line: {
					width: 3,
					gradient: {
						kind: 'linear',
						angle: 45,
						rotateWithShape: false,
						stops: [
							{ position: 0, color: 'F0C808' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			}
		),
		issue: 'https://github.com/shbernal/pptx-ts/issues/56',
	},
	{
		...text(
			'addText.line.gradient.center',
			{
				line: {
					width: 3,
					gradient: {
						kind: 'radial',
						center: { x: 20, y: 70 },
						stops: [
							{ position: 0, color: 'F0C808' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			},
			['addText.line.gradient.center.x', 'addText.line.gradient.center.y']
		),
		issue: 'https://github.com/shbernal/pptx-ts/issues/56',
	},
	{
		...text(
			'addText.line.pattern',
			{ line: { width: 3, pattern: { preset: 'dkDnDiag', fgColor: '250F6B', bgColor: 'F0C808' } } },
			['addText.line.pattern.preset', 'addText.line.pattern.fgColor', 'addText.line.pattern.bgColor']
		),
		issue: 'https://github.com/shbernal/pptx-ts/issues/56',
	},
	{
		...text('addText.objectLock.noAdjustHandles', { objectLock: { noAdjustHandles: true } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/57',
	},
	{
		...text('addText.objectLock.noChangeArrowheads', { objectLock: { noChangeArrowheads: true } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/57',
	},
	{
		...text('addText.objectLock.noChangeAspect', { objectLock: { noChangeAspect: true } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/57',
	},
	{
		...text('addText.objectLock.noChangeShapeType', { objectLock: { noChangeShapeType: true } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/57',
	},
	{
		...text('addText.objectLock.noCrop', { objectLock: { noCrop: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'the writer warns that noCrop is not valid on a:spLocks and drops it: it is a picture lock (a:picLocks), which a text box does not have',
		},
	},
	{
		...text('addText.objectLock.noDrilldown', { objectLock: { noDrilldown: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'the writer warns that noDrilldown is not valid on a:spLocks and drops it: it is a graphic-frame/group lock, which a text box does not have',
		},
	},
	{
		...text('addText.objectLock.noEditPoints', { objectLock: { noEditPoints: true } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/57',
	},
	{
		...text('addText.objectLock.noGrp', { objectLock: { noGrp: true } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/57',
	},
	{
		...text('addText.objectLock.noMove', { objectLock: { noMove: true } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/57',
	},
	{
		...text('addText.objectLock.noResize', { objectLock: { noResize: true } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/57',
	},
	{
		...text('addText.objectLock.noRot', { objectLock: { noRot: true } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/57',
	},
	{
		...text('addText.objectLock.noSelect', { objectLock: { noSelect: true } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/57',
	},
	{
		...text('addText.objectLock.noTextEdit', { objectLock: { noTextEdit: true } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/57',
	},
	text('addText.objectName', { objectName: 'Renamed probe' }),
	{
		...text('addText.outline', { outline: { color: '9B1B30', size: 1.5 } }, [
			'addText.outline.color',
			'addText.outline.size',
		]),
		issue: 'https://github.com/shbernal/pptx-ts/issues/54',
	},
	text('addText.paraIndent', { paraIndent: 12 }),
	text('addText.paraMarginLeft', { paraMarginLeft: 18 }),
	text('addText.paraSpaceAfter', { paraSpaceAfter: 12 }),
	text('addText.paraSpaceBefore', { paraSpaceBefore: 12 }),
	text('addText.lineSpacing', { lineSpacing: 30 }),
	text(
		'addText.points',
		{ shape: 'custGeom', points: [{ x: 0, y: 0 }, { x: 5, y: 0 }, { x: 5, y: 1.5 }, { close: true }] },
		['addText.points.x', 'addText.points.y', 'addText.points.close']
	),
	textOn(
		'addText.points.moveTo',
		{
			shape: 'custGeom',
			points: [
				{ x: 0, y: 0 },
				{ x: 5, y: 1.5 },
			],
		},
		{
			points: [
				{ x: 0, y: 0 },
				{ x: 5, y: 1.5 },
				{ x: 0, y: 1.5, moveTo: true },
				{ x: 5, y: 0 },
			],
		}
	),
	textOn(
		'addText.points.curve',
		{
			shape: 'custGeom',
			points: [
				{ x: 0, y: 0 },
				{ x: 5, y: 0 },
			],
		},
		{ points: [{ x: 0, y: 0 }, { x: 5, y: 0 }, { curve: { type: 'arc', hR: 0.75, wR: 2.5, stAng: 0, swAng: 180 } }] },
		[
			'addText.points.curve.type',
			'addText.points.curve.hR',
			'addText.points.curve.wR',
			'addText.points.curve.stAng',
			'addText.points.curve.swAng',
		]
	),
	textOn('addText.rectRadius', { shape: 'roundRect' }, { rectRadius: 0.3 }),
	text('addText.rotate', { rotate: 30 }),
	{ ...text('addText.rtlMode', { rtlMode: true }), issue: 'https://github.com/shbernal/pptx-ts/issues/54' },
	text(
		'addText.shadow.outer',
		{ shadow: { type: 'outer', angle: 60, blur: 6, color: '250F6B', offset: 4, transparency: 40 } },
		[
			'addText.shadow.type',
			'addText.shadow.angle',
			'addText.shadow.blur',
			'addText.shadow.color',
			'addText.shadow.offset',
			'addText.shadow.transparency',
		]
	),
	{
		...textOn(
			'addText.shadow.rotateWithShape',
			{ shadow: { type: 'outer', blur: 6 } },
			{ shadow: { type: 'outer', blur: 6, rotateWithShape: true } }
		),
		issue: 'https://github.com/shbernal/pptx-ts/issues/56',
	},
	text('addText.shape:ellipse', { shape: 'ellipse' }),
	textOn('addText.shapeAdjust', { shape: 'roundRect' }, { shapeAdjust: { name: 'adj', value: 0.4 } }, [
		'addText.shapeAdjust.name',
		'addText.shapeAdjust.value',
	]),
	{
		...run('addText.softBreakBefore', 1, { softBreakBefore: true }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/54',
	},
	text('addText.strike', { strike: true }),
	{ ...text('addText.subscript', { subscript: true }), issue: 'https://github.com/shbernal/pptx-ts/issues/55' },
	{ ...text('addText.superscript', { superscript: true }), issue: 'https://github.com/shbernal/pptx-ts/issues/55' },
	{
		...text('addText.tabStops', { tabStops: [{ position: 1.5, alignment: 'r' }] }, [
			'addText.tabStops.position',
			'addText.tabStops.alignment',
		]),
		issue: 'https://github.com/shbernal/pptx-ts/issues/54',
	},
	text('addText.textDirection:vert', { textDirection: 'vert' }),
	{ ...text('addText.textWarp', { textWarp: 'textArchUp' }), issue: 'https://github.com/shbernal/pptx-ts/issues/54' },
	{ ...text('addText.transparency', { transparency: 40 }), issue: 'https://github.com/shbernal/pptx-ts/issues/54' },
	{
		...textOn(
			'addText.underline.color',
			{ underline: { style: 'sng' } },
			{ underline: { style: 'sng', color: '9B1B30' } }
		),
		issue: 'https://github.com/shbernal/pptx-ts/issues/54',
	},
	text('addText.vert:eaVert', { vert: 'eaVert' }),
	text('addText.wrap', { wrap: false }),
	{ ...text('addText.placeholder', { placeholder: 'body' }), issue: 'https://github.com/shbernal/pptx-ts/issues/61' },
	{
		...shape('addShape.align:center', { align: 'center' }),
		accepted: {
			kind: 'inert',
			reason:
				'addShape writes no text body, so a paragraph alignment has nothing to attach to and the writer emits nothing',
		},
	},
	shape('addShape.altText', { altText: 'A probe description' }),
	shapeOn('addShape.angleRange', 'blockArc', {}, { angleRange: [30, 200] }),
	shapeOn('addShape.arcThicknessRatio', 'blockArc', { angleRange: [30, 200] }, { arcThicknessRatio: 0.3 }),
	shape('addShape.fill.color', { fill: { color: 'F0C808' } }, ['addShape.fill', 'addShape.fill.color']),
	{
		...shape('addShape.fill.transparency', { fill: { color: 'DDE3F0', transparency: 40 } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/56',
	},
	shape('addShape.fill.type:none', { fill: { type: 'none' } }, ['addShape.fill.type']),
	shape(
		'addShape.fill.gradient.linear',
		{
			fill: {
				gradient: {
					kind: 'linear',
					angle: 45,
					stops: [
						{ position: 0, color: 'F0C808' },
						{ position: 100, color: '250F6B', transparency: 30 },
					],
				},
			},
		},
		[
			'addShape.fill.gradient.kind',
			'addShape.fill.gradient.angle',
			'addShape.fill.gradient.stops.color',
			'addShape.fill.gradient.stops.position',
			'addShape.fill.gradient.stops.transparency',
		]
	),
	{
		...shapeOn(
			'addShape.fill.gradient.scaled',
			'rect',
			{
				fill: {
					gradient: {
						kind: 'linear',
						angle: 45,
						stops: [
							{ position: 0, color: 'F0C808' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			},
			{
				fill: {
					gradient: {
						kind: 'linear',
						angle: 45,
						scaled: true,
						stops: [
							{ position: 0, color: 'F0C808' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			}
		),
		issue: 'https://github.com/shbernal/pptx-ts/issues/56',
	},
	{
		...shapeOn(
			'addShape.fill.gradient.rotateWithShape',
			'rect',
			{
				fill: {
					gradient: {
						kind: 'linear',
						angle: 45,
						stops: [
							{ position: 0, color: 'F0C808' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			},
			{
				fill: {
					gradient: {
						kind: 'linear',
						angle: 45,
						rotateWithShape: false,
						stops: [
							{ position: 0, color: 'F0C808' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			}
		),
		issue: 'https://github.com/shbernal/pptx-ts/issues/56',
	},
	{
		...shape(
			'addShape.fill.gradient.center',
			{
				fill: {
					gradient: {
						kind: 'radial',
						center: { x: 20, y: 70 },
						stops: [
							{ position: 0, color: 'F0C808' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			},
			['addShape.fill.gradient.center.x', 'addShape.fill.gradient.center.y']
		),
		issue: 'https://github.com/shbernal/pptx-ts/issues/56',
	},
	shape('addShape.fill.pattern', { fill: { pattern: { preset: 'dkDnDiag', fgColor: '250F6B', bgColor: 'F0C808' } } }, [
		'addShape.fill.pattern.preset',
		'addShape.fill.pattern.fgColor',
		'addShape.fill.pattern.bgColor',
	]),
	shape('addShape.fill.image.data', { fill: { image: { data: PNG } } }),
	shapeOn(
		'addShape.fill.image.crop',
		'rect',
		{ fill: { image: { data: PNG } } },
		{ fill: { image: { data: PNG, crop: { l: 10, t: 20, r: 30, b: 5 } } } },
		[
			'addShape.fill.image.crop.l',
			'addShape.fill.image.crop.t',
			'addShape.fill.image.crop.r',
			'addShape.fill.image.crop.b',
		]
	),
	shape('addShape.flipH', { flipH: true }),
	shape('addShape.flipV', { flipV: true }),
	shape('addShape.x', { x: 2 }),
	shape('addShape.y', { y: 2 }),
	shape('addShape.w', { w: 6 }),
	shape('addShape.h', { h: 2 }),
	shape('addShape.hyperlink.url', { hyperlink: { url: 'https://example.com/', tooltip: 'Example' } }, [
		'addShape.hyperlink.url',
		'addShape.hyperlink.tooltip',
	]),
	shape('addShape.hyperlink.slide', { hyperlink: { slide: 1 } }),
	shape('addShape.hyperlink.action:nextslide', { hyperlink: { action: 'nextslide' } }, ['addShape.hyperlink.action']),
	shape('addShape.line', { line: { color: '250F6B', width: 2 } }, ['addShape.line.color', 'addShape.line.width']),
	shapeOn(
		'addShape.line.cap:round',
		'rect',
		{ line: { color: '250F6B', width: 2 } },
		{ line: { color: '250F6B', width: 2, cap: 'round' } }
	),
	shapeOn(
		'addShape.line.beginArrowType',
		'rect',
		{ line: { color: '250F6B', width: 2 } },
		{ line: { color: '250F6B', width: 2, beginArrowType: 'triangle' } }
	),
	shapeOn(
		'addShape.line.endArrowType',
		'rect',
		{ line: { color: '250F6B', width: 2 } },
		{ line: { color: '250F6B', width: 2, endArrowType: 'oval' } }
	),
	shapeOn(
		'addShape.line.transparency',
		'rect',
		{ line: { color: '250F6B', width: 2 } },
		{ line: { color: '250F6B', width: 2, transparency: 40 } }
	),
	shapeOn(
		'addShape.line.type:none',
		'rect',
		{ line: { color: '250F6B', width: 2 } },
		{ line: { color: '250F6B', width: 2, type: 'none' } }
	),
	{
		...shape(
			'addShape.line.gradient.linear',
			{
				line: {
					width: 3,
					gradient: {
						kind: 'linear',
						angle: 45,
						stops: [
							{ position: 0, color: 'F0C808' },
							{ position: 100, color: '250F6B', transparency: 30 },
						],
					},
				},
			},
			[
				'addShape.line.gradient.kind',
				'addShape.line.gradient.angle',
				'addShape.line.gradient.stops.color',
				'addShape.line.gradient.stops.position',
				'addShape.line.gradient.stops.transparency',
			]
		),
		issue: 'https://github.com/shbernal/pptx-ts/issues/56',
	},
	{
		...shapeOn(
			'addShape.line.gradient.scaled',
			'rect',
			{
				line: {
					width: 3,
					gradient: {
						kind: 'linear',
						angle: 45,
						stops: [
							{ position: 0, color: 'F0C808' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			},
			{
				line: {
					width: 3,
					gradient: {
						kind: 'linear',
						angle: 45,
						scaled: true,
						stops: [
							{ position: 0, color: 'F0C808' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			}
		),
		issue: 'https://github.com/shbernal/pptx-ts/issues/56',
	},
	{
		...shapeOn(
			'addShape.line.gradient.rotateWithShape',
			'rect',
			{
				line: {
					width: 3,
					gradient: {
						kind: 'linear',
						angle: 45,
						stops: [
							{ position: 0, color: 'F0C808' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			},
			{
				line: {
					width: 3,
					gradient: {
						kind: 'linear',
						angle: 45,
						rotateWithShape: false,
						stops: [
							{ position: 0, color: 'F0C808' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			}
		),
		issue: 'https://github.com/shbernal/pptx-ts/issues/56',
	},
	{
		...shape(
			'addShape.line.gradient.center',
			{
				line: {
					width: 3,
					gradient: {
						kind: 'radial',
						center: { x: 20, y: 70 },
						stops: [
							{ position: 0, color: 'F0C808' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			},
			['addShape.line.gradient.center.x', 'addShape.line.gradient.center.y']
		),
		issue: 'https://github.com/shbernal/pptx-ts/issues/56',
	},
	{
		...shape(
			'addShape.line.pattern',
			{ line: { width: 3, pattern: { preset: 'dkDnDiag', fgColor: '250F6B', bgColor: 'F0C808' } } },
			['addShape.line.pattern.preset', 'addShape.line.pattern.fgColor', 'addShape.line.pattern.bgColor']
		),
		issue: 'https://github.com/shbernal/pptx-ts/issues/56',
	},
	{
		...shape('addShape.objectLock.noAdjustHandles', { objectLock: { noAdjustHandles: true } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/57',
	},
	{
		...shape('addShape.objectLock.noChangeArrowheads', { objectLock: { noChangeArrowheads: true } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/57',
	},
	{
		...shape('addShape.objectLock.noChangeAspect', { objectLock: { noChangeAspect: true } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/57',
	},
	{
		...shape('addShape.objectLock.noChangeShapeType', { objectLock: { noChangeShapeType: true } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/57',
	},
	{
		...shape('addShape.objectLock.noCrop', { objectLock: { noCrop: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'the writer warns that noCrop is not valid on a:spLocks and drops it: it is a picture lock (a:picLocks), which a shape does not have',
		},
	},
	{
		...shape('addShape.objectLock.noDrilldown', { objectLock: { noDrilldown: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'the writer warns that noDrilldown is not valid on a:spLocks and drops it: it is a graphic-frame/group lock, which a shape does not have',
		},
	},
	{
		...shape('addShape.objectLock.noEditPoints', { objectLock: { noEditPoints: true } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/57',
	},
	{
		...shape('addShape.objectLock.noGrp', { objectLock: { noGrp: true } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/57',
	},
	{
		...shape('addShape.objectLock.noMove', { objectLock: { noMove: true } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/57',
	},
	{
		...shape('addShape.objectLock.noResize', { objectLock: { noResize: true } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/57',
	},
	{
		...shape('addShape.objectLock.noRot', { objectLock: { noRot: true } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/57',
	},
	{
		...shape('addShape.objectLock.noSelect', { objectLock: { noSelect: true } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/57',
	},
	{
		...shape('addShape.objectLock.noTextEdit', { objectLock: { noTextEdit: true } }),
		issue: 'https://github.com/shbernal/pptx-ts/issues/57',
	},
	shape('addShape.objectName', { objectName: 'Renamed probe' }),
	shapeOn(
		'addShape.points',
		'custGeom',
		{
			points: [
				{ x: 0, y: 0 },
				{ x: 5, y: 0 },
			],
		},
		{ points: [{ x: 0, y: 0 }, { x: 5, y: 0 }, { x: 5, y: 1.5 }, { close: true }] },
		['addShape.points.x', 'addShape.points.y', 'addShape.points.close']
	),
	shapeOn(
		'addShape.points.moveTo',
		'custGeom',
		{
			points: [
				{ x: 0, y: 0 },
				{ x: 5, y: 1.5 },
			],
		},
		{
			points: [
				{ x: 0, y: 0 },
				{ x: 5, y: 1.5 },
				{ x: 0, y: 1.5, moveTo: true },
				{ x: 5, y: 0 },
			],
		}
	),
	shapeOn(
		'addShape.points.curve',
		'custGeom',
		{
			points: [
				{ x: 0, y: 0 },
				{ x: 5, y: 0 },
			],
		},
		{ points: [{ x: 0, y: 0 }, { x: 5, y: 0 }, { curve: { type: 'arc', hR: 0.75, wR: 2.5, stAng: 0, swAng: 180 } }] },
		[
			'addShape.points.curve.type',
			'addShape.points.curve.hR',
			'addShape.points.curve.wR',
			'addShape.points.curve.stAng',
			'addShape.points.curve.swAng',
		]
	),
	{
		...shapeOn('addShape.guides', 'custGeom', { points: CUST }, { guides: [{ name: 'g0', formula: 'val 10800' }] }, [
			'addShape.guides.name',
			'addShape.guides.formula',
		]),
		issue: 'https://github.com/shbernal/pptx-ts/issues/78',
	},
	{
		...shapeOn(
			'addShape.connectionSites',
			'custGeom',
			{ points: CUST },
			{ connectionSites: [{ ang: 90, x: 2.5, y: 1.5 }] },
			['addShape.connectionSites.ang', 'addShape.connectionSites.x', 'addShape.connectionSites.y']
		),
		issue: 'https://github.com/shbernal/pptx-ts/issues/78',
	},
	{
		...shapeOn(
			'addShape.adjustHandles.xy',
			'custGeom',
			{
				points: CUST,
				guides: [
					{ name: 'g0', formula: 'val 10800' },
					{ name: 'g1', formula: 'val 5400' },
				],
			},
			{ adjustHandles: [{ x: 1, y: 0.5, gdRefX: 'g0', minX: 0, maxX: 5, gdRefY: 'g1', minY: 0, maxY: 1.5 }] },
			[
				'addShape.adjustHandles.x',
				'addShape.adjustHandles.y',
				'addShape.adjustHandles.gdRefX',
				'addShape.adjustHandles.minX',
				'addShape.adjustHandles.maxX',
				'addShape.adjustHandles.gdRefY',
				'addShape.adjustHandles.minY',
				'addShape.adjustHandles.maxY',
			]
		),
		issue: 'https://github.com/shbernal/pptx-ts/issues/78',
	},
	{
		...shapeOn(
			'addShape.adjustHandles.polar',
			'custGeom',
			{
				points: CUST,
				guides: [
					{ name: 'g0', formula: 'val 10800' },
					{ name: 'g1', formula: 'val 5400000' },
				],
			},
			{ adjustHandles: [{ x: 2.5, y: 0.75, gdRefR: 'g0', minR: 0, maxR: 1, gdRefAng: 'g1', minAng: 0, maxAng: 180 }] },
			[
				'addShape.adjustHandles.gdRefR',
				'addShape.adjustHandles.minR',
				'addShape.adjustHandles.maxR',
				'addShape.adjustHandles.gdRefAng',
				'addShape.adjustHandles.minAng',
				'addShape.adjustHandles.maxAng',
			]
		),
		issue: 'https://github.com/shbernal/pptx-ts/issues/78',
	},
	shapeOn('addShape.rectRadius', 'roundRect', {}, { rectRadius: 0.3 }),
	shape(
		'addShape.shadow.outer',
		{ shadow: { type: 'outer', angle: 60, blur: 6, color: '250F6B', offset: 4, transparency: 40 } },
		[
			'addShape.shadow.type',
			'addShape.shadow.angle',
			'addShape.shadow.blur',
			'addShape.shadow.color',
			'addShape.shadow.offset',
			'addShape.shadow.transparency',
		]
	),
	{
		...shapeOn(
			'addShape.shadow.rotateWithShape',
			'rect',
			{ shadow: { type: 'outer', blur: 6 } },
			{ shadow: { type: 'outer', blur: 6, rotateWithShape: true } }
		),
		issue: 'https://github.com/shbernal/pptx-ts/issues/56',
	},
	shapeOn('addShape.shapeAdjust', 'roundRect', {}, { shapeAdjust: { name: 'adj', value: 0.4 } }, [
		'addShape.shapeAdjust.name',
		'addShape.shapeAdjust.value',
	]),
]

/**
 * Inventory paths deliberately not probed, and why. A path here counts as covered,
 * so the reason has to hold up: "hard to probe" is not one; "selects the input
 * source and writes nothing of its own" is.
 */
export const SKIPPED: Record<string, string> = {
	'addText.data': 'an image source for picture-filled text, exercised by addText.fill.image.*',
	'addText.path': 'an image source for picture-filled text, exercised by addText.fill.image.*',
	'addText.fill.image.path':
		'the file-path spelling of the picture source; the picture it writes is exercised by addText.fill.image.data',
	'addText.bullet.image.path':
		'the file-path spelling of the bullet picture source; the bullet it writes is exercised by addText.bullet.image.data',
	'addShape.fill.image.path':
		'the file-path spelling of the picture source; the picture it writes is exercised by addShape.fill.image.data',
}
