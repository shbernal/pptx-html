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
import type {
	ConnectorProps,
	GroupProps,
	ImageProps,
	MasterTextStyleLevel,
	OptsChartData,
	SHAPE_NAME,
	ShapeProps,
	Slide,
	SlideMasterProps,
	SlideNumberProps,
	TableCellProps,
	TableProps,
	NotesTextOptions,
	TextPropsOptions,
	ThemeProps,
} from 'pptx-ts'
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

/** Speaker notes on a plain slide, one run; `options` go on the run only when on. */
function notes(id: string, options: NotesTextOptions, covers = [id.split(':')[0] as string]): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addNotes([{ text: 'Probe notes', options: on ? options : {} }])
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

/** A 2×2 table with `base` on the first cell on both sides, for options that only mean something on top of another. */
function cellOn(
	id: string,
	base: TableCellProps,
	options: TableCellProps,
	covers = [id.split(':')[0] as string]
): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addTable(
				[
					[{ text: 'a', options: { ...base, ...(on ? options : {}) } }, { text: 'b' }],
					[{ text: 'c' }, { text: 'd' }],
				],
				{ ...BOX, objectName: 'probe' }
			)
		},
	}
}

/** A 2×2 table whose first cell holds two runs; `options` go on run `at` only. */
function cellRun(id: string, at: 0 | 1, options: TableCellProps, covers = [id.split(':')[0] as string]): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addTable(
				[
					[
						{ text: ['First run', 'Second run'].map((text, i) => ({ text, options: on && i === at ? options : {} })) },
						{ text: 'b' },
					],
					[{ text: 'c' }, { text: 'd' }],
				],
				{ ...BOX, objectName: 'probe' }
			)
		},
	}
}

/**
 * A 2×2 table whose first cell spans into its neighbour when on. Off, the cell it
 * would cover is present and empty, so the two sides differ by the merge alone.
 */
function span(
	id: string,
	options: { colspan: number } | { rowspan: number },
	covers = [id.split(':')[0] as string]
): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			const across = 'colspan' in options
			const first = { text: 'a', options: on ? options : {} }
			pptx
				.addSlide()
				.addTable(
					across
						? [on ? [first] : [first, { text: '' }], [{ text: 'c' }, { text: 'd' }]]
						: [[first, { text: 'b' }], on ? [{ text: 'd' }] : [{ text: '' }, { text: 'd' }]],
					{ ...BOX, objectName: 'probe' }
				)
		},
	}
}

/** A slide holding the probe text box; `apply` sets the slide's own properties on the probe side only. */
function slide(id: string, apply: (slide: Slide) => void, covers = [id.split(':')[0] as string]): Probe {
	return slideOn(id, () => {}, apply, covers)
}

/** A slide holding the probe text box, with `base` on both sides and `apply` on the probe side only. */
function slideOn(
	id: string,
	base: (slide: Slide) => void,
	apply: (slide: Slide) => void,
	covers = [id.split(':')[0] as string]
): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			const target = pptx.addSlide()
			target.addText('Probe text', { ...BOX, objectName: 'probe' })
			base(target)
			if (on) apply(target)
		},
	}
}

const NUMBER_BOX = { x: 0.5, y: 6.5, w: 1, h: 0.4 }

/** A slide number in a fixed box, with `base` on both sides and `options` merged in on the probe side. */
function slideNumber(
	id: string,
	options: SlideNumberProps,
	base: SlideNumberProps = {},
	covers = [id.split(':')[0] as string]
): Probe {
	return slideOn(
		id,
		(target) => {
			target.slideNumber = { ...NUMBER_BOX, ...base }
		},
		(target) => {
			target.slideNumber = { ...NUMBER_BOX, ...base, ...options }
		},
		covers
	)
}

/** Master options a probe sets; `title` is fixed by the helper so the slide can name it. */
type MasterOptions = Omit<SlideMasterProps, 'title'>

/** A master the probe slide uses, with `base` on both sides and `options` merged in on the probe side. */
function master(
	id: string,
	options: MasterOptions,
	base: MasterOptions = {},
	covers = [id.split(':')[0] as string]
): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			pptx.defineSlideMaster({ title: 'PROBE', ...base, ...(on ? options : {}) })
			pptx.addSlide({ masterTitle: 'PROBE' }).addText('Probe text', { ...BOX, objectName: 'probe' })
		},
	}
}

/** A linear gradient both sides of a gradient-knob probe share. */
const LINEAR = {
	kind: 'linear' as const,
	angle: 45,
	stops: [
		{ position: 0, color: 'F0C808' },
		{ position: 100, color: '250F6B' },
	],
}

/** The chart box and series a master chart probe places, so only the series option differs. */
const CHART_BOX = { x: 1, y: 3, w: 5, h: 3 }
const SERIES = { name: 'Sales', labels: ['A', 'B', 'C'], values: [3, 5, 4] }

/** A master bar chart whose one series takes `series` on the probe side. */
function masterChart(id: string, series: Partial<OptsChartData>, covers = [id.split(':')[0] as string]): Probe {
	return master(
		id,
		{ objects: [{ chart: { type: 'bar', data: [{ ...SERIES, ...series }], options: CHART_BOX } }] },
		{ objects: [{ chart: { type: 'bar', data: [SERIES], options: CHART_BOX } }] },
		covers
	)
}

/** The two series of a bubble chart: x values, then y values that take the sizes. */
const BUBBLE_X = { name: 'X', values: [1, 2, 3] }
const BUBBLE_Y = { name: 'Y', values: [4, 6, 5] }

/** A bar-and-line combo on two axis pairs, the master chart's multi-type form. */
const COMBO = [
	{ type: 'bar' as const, data: [SERIES], options: {} },
	{
		type: 'line' as const,
		data: [{ ...SERIES, name: 'Trend' }],
		options: { secondaryValAxis: true, secondaryCatAxis: true },
	},
]
const COMBO_BOX = { ...CHART_BOX, valAxes: [{}, {}], catAxes: [{}, {}] }

/** A master slide number in a fixed box, with `base` on both sides and `options` merged in on the probe side. */
function masterNumber(
	id: string,
	options: SlideNumberProps,
	base: SlideNumberProps = {},
	covers = [id.split(':')[0] as string]
): Probe {
	return master(
		id,
		{ slideNumber: { ...NUMBER_BOX, ...base, ...options } },
		{ slideNumber: { ...NUMBER_BOX, ...base } },
		covers
	)
}

/**
 * Master text styles for `group` with `base` on both sides and `options` merged in on the probe side.
 *
 * `title` and `body` style only the text of a placeholder of that type, so the master
 * declares one and the slide fills it: without it the style is written and nothing on
 * the slide inherits it, which round-trips and is never rendered. `other` has no slide
 * text to reach (a slide text box takes the presentation's default style, not the
 * master's), so its probe keeps the plain text box and is not rendered by design.
 *
 * The slide text states `'inherit'` for bullet, margin and indent, or the writer's own
 * `a:buNone` and zero insets would override the master's. Even so, the paragraph-level
 * styles (align, margins, bullets) are not rendered yet: the read model resolves a run's
 * inherited properties and none of a paragraph's (ts-pptx#143:
 * https://github.com/shbernal/pptx-ts/issues/143). They render, and the ledger moves,
 * when `Paragraph` gains resolved accessors and `src/import/text.ts` reads them.
 */
function masterStyle(
	id: string,
	group: 'title' | 'body' | 'other',
	options: MasterTextStyleLevel,
	base: MasterTextStyleLevel = {},
	covers = [id.split(':')[0] as string]
): Probe {
	const styles = (level: MasterTextStyleLevel) => ({
		textStyles: group === 'title' ? { title: level } : { [group]: [level] },
	})
	if (group === 'other') return master(id, styles({ ...base, ...options }), styles(base), covers)
	const placeholder = { name: group, type: group, ...BOX }
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			pptx.defineSlideMaster({
				title: 'PROBE',
				objects: [{ placeholder: { options: placeholder, text: '' } }],
				...styles(on ? { ...base, ...options } : base),
			})
			pptx.addSlide({ masterTitle: 'PROBE' }).addText('Probe text', {
				placeholder: group,
				objectName: 'probe',
				bullet: 'inherit',
				paraMarginLeft: 'inherit',
				paraIndent: 'inherit',
			})
		},
	}
}

/** A silent eight-sample WAV, for transition sounds. */
const WAV = 'audio/wav;base64,UklGRiwAAABXQVZFZm10IBAAAAABAAEAQB8AAEAfAAABAAgAZGF0YQgAAACAgICAgICAgA=='

/** `addImage` of the one-pixel PNG in the probe box, with `base` on both sides and `options` on the probe side. */
function imageOn(
	id: string,
	base: Omit<ImageProps, 'data' | 'path'>,
	options: Omit<ImageProps, 'data' | 'path'>,
	covers = [id.split(':')[0] as string]
): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addImage({ data: PNG, ...BOX, objectName: 'probe', ...base, ...(on ? options : {}) })
		},
	}
}

/** `addImage` of the one-pixel PNG in the probe box; `options` is applied only when the probe is on. */
function image(id: string, options: Omit<ImageProps, 'data' | 'path'>, covers = [id.split(':')[0] as string]): Probe {
	return imageOn(id, {}, options, covers)
}

/** The probe connector's endpoints, on both sides. */
const LINE = { x1: 1, y1: 1, x2: 5, y2: 3 }

/**
 * `addConnector` between two named shapes, so endpoint bindings have targets on
 * both sides. `base` is on both sides; `options` on the probe side only.
 */
function connector(
	id: string,
	base: Partial<ConnectorProps>,
	options: Partial<ConnectorProps>,
	covers = [id.split(':')[0] as string]
): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			const target = pptx.addSlide()
			target.addShape('rect', { x: 0.5, y: 0.5, w: 1, h: 1, fill: { color: 'DDE3F0' }, objectName: 'from' })
			target.addShape('rect', { x: 4.5, y: 2.5, w: 1, h: 1, fill: { color: 'DDE3F0' }, objectName: 'to' })
			target.addConnector({ ...LINE, objectName: 'probe', ...base, ...(on ? options : {}) })
		},
	}
}

/** A lock flag with no attribute on `a:grpSpLocks`, which the writer drops on a group. */
const NOT_A_GROUP_LOCK = {
	kind: 'inert',
	reason:
		'not an attribute of a:grpSpLocks: CT_GroupLocking carries only noGrp, noUngrp, noSelect, noRot, noChangeAspect, noMove and noResize, so a group has nowhere to write it',
} as const

/** The probe group's explicit frame: the bounding box of its two children. */
const GROUP_FRAME = { x: 1, y: 1, w: 4, h: 1 }

/**
 * `addGroup` of two filled rects. `base` is on both sides; `options` on the probe
 * side only. The frame is all-or-nothing, so a frame probe sets `GROUP_FRAME` as
 * its base and moves one edge.
 */
function group(id: string, base: GroupProps, options: GroupProps, covers = [id.split(':')[0] as string]): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			pptx
				.addSlide()
				.addGroup(
					[
						{ rect: { x: 1, y: 1, w: 1, h: 1, fill: { color: 'DDE3F0' } } },
						{ rect: { x: 4, y: 1, w: 1, h: 1, fill: { color: '250F6B' } } },
					],
					{ objectName: 'probe', ...base, ...(on ? options : {}) }
				)
		},
	}
}

/** `pptx.theme` set on the probe side only, over a slide holding the probe text box. */
function theme(id: string, options: ThemeProps, covers = [id.split(':')[0] as string]): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			if (on) pptx.theme = options
			pptx.addSlide().addText('Probe text', { ...BOX, objectName: 'probe' })
		},
	}
}

export const PROBES: Probe[] = [
	text('addText.bold', { bold: true }),
	text('addText.italic', { italic: true }),
	text('addText.underline.style:dbl', { underline: { style: 'dbl' } }),
	text('addText.strike:dblStrike', { strike: 'dblStrike' }),
	text('addText.charSpacing', { charSpacing: 4 }),
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
	text('addText.baseline', { baseline: 30 }),
	run('addText.breakLine', 0, { breakLine: true }),
	text('addText.bullet', { bullet: true }),
	textOn('addText.bullet.characterCode', { bullet: true }, { bullet: { characterCode: '25BA' } }),
	textOn('addText.bullet.color', { bullet: true }, { bullet: { color: '9B1B30' } }),
	textOn('addText.bullet.fontFace', { bullet: true }, { bullet: { fontFace: 'Wingdings' } }),
	textOn('addText.bullet.image.data', { bullet: true }, { bullet: { image: { data: PNG } } }),
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
	text('addText.columns', { columns: 2 }),
	textOn('addText.columnSpacing', { columns: 2 }, { columnSpacing: 0.5 }),
	text('addText.fill.color', { fill: { color: 'F0C808' } }, ['addText.fill', 'addText.fill.color']),
	textOn('addText.fill.transparency', { fill: { color: 'F0C808' } }, { fill: { color: 'F0C808', transparency: 40 } }),
	text('addText.fill.type:inherit', { fill: { type: 'inherit' } }),
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
	textOn(
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
	textOn(
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
	text(
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
	textOn('addText.fontFaceEA', { fontFace: 'Georgia' }, { fontFaceEA: 'MS Mincho' }),
	text('addText.fontSize', { fontSize: 28 }),
	text('addText.glow', { glow: { size: 8, opacity: 0.5, color: 'FF00FF' } }, [
		'addText.glow.size',
		'addText.glow.opacity',
		'addText.glow.color',
	]),
	text('addText.x', { x: 2 }),
	text('addText.y', { y: 2 }),
	text('addText.w', { w: 6 }),
	text('addText.h', { h: 2 }),
	text('addText.highlight', { highlight: 'FFFF00' }),
	text('addText.hyperlink.url', { hyperlink: { url: 'https://example.com/', tooltip: 'Example' } }, [
		'addText.hyperlink.url',
		'addText.hyperlink.tooltip',
	]),
	text('addText.hyperlink.slide', { hyperlink: { slide: 1 } }),
	text('addText.hyperlink.action:nextslide', { hyperlink: { action: 'nextslide' } }),
	text('addText.indentLevel', { indentLevel: 2 }),
	text('addText.isTextBox', { isTextBox: true }),
	text('addText.lang', { lang: 'fr-FR' }),
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
	text(
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
	textOn(
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
	textOn(
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
	text(
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
	text(
		'addText.line.pattern',
		{ line: { width: 3, pattern: { preset: 'dkDnDiag', fgColor: '250F6B', bgColor: 'F0C808' } } },
		['addText.line.pattern.preset', 'addText.line.pattern.fgColor', 'addText.line.pattern.bgColor']
	),
	text('addText.objectLock.noAdjustHandles', { objectLock: { noAdjustHandles: true } }),
	text('addText.objectLock.noChangeArrowheads', { objectLock: { noChangeArrowheads: true } }),
	text('addText.objectLock.noChangeAspect', { objectLock: { noChangeAspect: true } }),
	text('addText.objectLock.noChangeShapeType', { objectLock: { noChangeShapeType: true } }),
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
	text('addText.objectLock.noEditPoints', { objectLock: { noEditPoints: true } }),
	text('addText.objectLock.noGrp', { objectLock: { noGrp: true } }),
	text('addText.objectLock.noMove', { objectLock: { noMove: true } }),
	text('addText.objectLock.noResize', { objectLock: { noResize: true } }),
	text('addText.objectLock.noRot', { objectLock: { noRot: true } }),
	text('addText.objectLock.noSelect', { objectLock: { noSelect: true } }),
	text('addText.objectLock.noTextEdit', { objectLock: { noTextEdit: true } }),
	text('addText.objectName', { objectName: 'Renamed probe' }),
	text('addText.outline', { outline: { color: '9B1B30', size: 1.5 } }, [
		'addText.outline.color',
		'addText.outline.size',
	]),
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
	text('addText.rtlMode', { rtlMode: true }),
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
	textOn(
		'addText.shadow.rotateWithShape',
		{ shadow: { type: 'outer', blur: 6 } },
		{ shadow: { type: 'outer', blur: 6, rotateWithShape: true } }
	),
	text('addText.shape:ellipse', { shape: 'ellipse' }),
	textOn('addText.shapeAdjust', { shape: 'roundRect' }, { shapeAdjust: { name: 'adj', value: 0.4 } }, [
		'addText.shapeAdjust.name',
		'addText.shapeAdjust.value',
	]),
	run('addText.softBreakBefore', 1, { softBreakBefore: true }),
	text('addText.strike', { strike: true }),
	text('addText.subscript', { subscript: true }),
	text('addText.superscript', { superscript: true }),
	text('addText.tabStops', { tabStops: [{ position: 1.5, alignment: 'r' }] }, [
		'addText.tabStops.position',
		'addText.tabStops.alignment',
	]),
	text('addText.textDirection:vert', { textDirection: 'vert' }),
	text('addText.textWarp', { textWarp: 'textArchUp' }),
	text('addText.transparency', { transparency: 40 }),
	textOn('addText.underline.color', { underline: { style: 'sng' } }, { underline: { style: 'sng', color: '9B1B30' } }),
	text('addText.vert:eaVert', { vert: 'eaVert' }),
	text('addText.wrap', { wrap: false }),
	text('addText.placeholder', { placeholder: 'body' }),
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
	shape('addShape.fill.transparency', { fill: { color: 'DDE3F0', transparency: 40 } }),
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
	shapeOn(
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
	shapeOn(
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
	shape(
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
	shape(
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
	shapeOn(
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
	shapeOn(
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
	shape(
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
	shape(
		'addShape.line.pattern',
		{ line: { width: 3, pattern: { preset: 'dkDnDiag', fgColor: '250F6B', bgColor: 'F0C808' } } },
		['addShape.line.pattern.preset', 'addShape.line.pattern.fgColor', 'addShape.line.pattern.bgColor']
	),
	shape('addShape.objectLock.noAdjustHandles', { objectLock: { noAdjustHandles: true } }),
	shape('addShape.objectLock.noChangeArrowheads', { objectLock: { noChangeArrowheads: true } }),
	shape('addShape.objectLock.noChangeAspect', { objectLock: { noChangeAspect: true } }),
	shape('addShape.objectLock.noChangeShapeType', { objectLock: { noChangeShapeType: true } }),
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
	shape('addShape.objectLock.noEditPoints', { objectLock: { noEditPoints: true } }),
	shape('addShape.objectLock.noGrp', { objectLock: { noGrp: true } }),
	shape('addShape.objectLock.noMove', { objectLock: { noMove: true } }),
	shape('addShape.objectLock.noResize', { objectLock: { noResize: true } }),
	shape('addShape.objectLock.noRot', { objectLock: { noRot: true } }),
	shape('addShape.objectLock.noSelect', { objectLock: { noSelect: true } }),
	shape('addShape.objectLock.noTextEdit', { objectLock: { noTextEdit: true } }),
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
	shapeOn(
		'addShape.shadow.rotateWithShape',
		'rect',
		{ shadow: { type: 'outer', blur: 6 } },
		{ shadow: { type: 'outer', blur: 6, rotateWithShape: true } }
	),
	shapeOn('addShape.shapeAdjust', 'roundRect', {}, { shapeAdjust: { name: 'adj', value: 0.4 } }, [
		'addShape.shapeAdjust.name',
		'addShape.shapeAdjust.value',
	]),
	cell('addTable.cell.align:center', { align: 'center' }),
	cell('addTable.cell.anchorCtr', { anchorCtr: true }),
	cell('addTable.cell.baseline', { baseline: 30 }),
	cell('addTable.cell.bold', { bold: true }),
	cell('addTable.cell.border', { border: { type: 'solid', color: '9B1B30', width: 2 } }, [
		'addTable.cell.border.type',
		'addTable.cell.border.color',
		'addTable.cell.border.width',
	]),
	cellOn(
		'addTable.cell.border.cap:round',
		{ border: { type: 'solid', color: '9B1B30', width: 2 } },
		{ border: { type: 'solid', color: '9B1B30', width: 2, cap: 'round' } }
	),
	cellOn(
		'addTable.cell.border.dashType:dash',
		{ border: { type: 'solid', color: '9B1B30', width: 2 } },
		{ border: { type: 'solid', color: '9B1B30', width: 2, dashType: 'dash' } }
	),
	cellOn(
		'addTable.cell.border.transparency',
		{ border: { type: 'solid', color: '9B1B30', width: 2 } },
		{ border: { type: 'solid', color: '9B1B30', width: 2, transparency: 40 } }
	),
	cellRun('addTable.cell.breakLine', 0, { breakLine: true }),
	cell('addTable.cell.bullet', { bullet: true }),
	cellOn('addTable.cell.bullet.characterCode', { bullet: true }, { bullet: { characterCode: '25BA' } }),
	cellOn('addTable.cell.bullet.color', { bullet: true }, { bullet: { color: '9B1B30' } }),
	cellOn('addTable.cell.bullet.fontFace', { bullet: true }, { bullet: { fontFace: 'Wingdings' } }),
	cellOn('addTable.cell.bullet.image.data', { bullet: true }, { bullet: { image: { data: PNG } } }),
	cellOn('addTable.cell.bullet.indent', { bullet: true }, { bullet: { indent: 10 } }),
	cellOn(
		'addTable.cell.bullet.numberStartAt',
		{ bullet: { type: 'number' } },
		{ bullet: { type: 'number', numberStartAt: 5 } }
	),
	cellOn(
		'addTable.cell.bullet.numberType',
		{ bullet: { type: 'number' } },
		{ bullet: { type: 'number', numberType: 'romanLcParenR' } }
	),
	cellOn('addTable.cell.bullet.size', { bullet: true }, { bullet: { size: 80 } }),
	cell('addTable.cell.bullet.type:number', { bullet: { type: 'number' } }),
	cell('addTable.cell.caps:all', { caps: 'all' }),
	cell('addTable.cell.cell3D', { cell3D: { preset: 'angle' } }, [
		'addTable.cell.cell3D',
		'addTable.cell.cell3D.preset',
	]),
	cellOn('addTable.cell.cell3D.height', { cell3D: { preset: 'angle' } }, { cell3D: { preset: 'angle', height: 3 } }),
	cellOn('addTable.cell.cell3D.width', { cell3D: { preset: 'angle' } }, { cell3D: { preset: 'angle', width: 4 } }),
	cellOn(
		'addTable.cell.cell3D.material:metal',
		{ cell3D: { preset: 'angle' } },
		{ cell3D: { preset: 'angle', material: 'metal' } }
	),
	cellOn(
		'addTable.cell.cell3D.lightRig',
		{ cell3D: { preset: 'angle' } },
		{ cell3D: { preset: 'angle', lightRig: { rig: 'threePt', dir: 't' } } },
		['addTable.cell.cell3D.lightRig.rig', 'addTable.cell.cell3D.lightRig.dir']
	),
	cell('addTable.cell.color', { color: '9B1B30' }),
	span('addTable.cell.colspan', { colspan: 2 }),
	span('addTable.cell.rowspan', { rowspan: 2 }),
	cell('addTable.cell.diagonal.tlToBr', { diagonal: { tlToBr: { type: 'solid', color: '9B1B30', width: 2 } } }, [
		'addTable.cell.diagonal.tlToBr.type',
		'addTable.cell.diagonal.tlToBr.color',
		'addTable.cell.diagonal.tlToBr.width',
	]),
	cellOn(
		'addTable.cell.diagonal.tlToBr.cap:round',
		{ diagonal: { tlToBr: { type: 'solid', color: '9B1B30', width: 2 } } },
		{ diagonal: { tlToBr: { type: 'solid', color: '9B1B30', width: 2, cap: 'round' } } }
	),
	cellOn(
		'addTable.cell.diagonal.tlToBr.dashType:dash',
		{ diagonal: { tlToBr: { type: 'solid', color: '9B1B30', width: 2 } } },
		{ diagonal: { tlToBr: { type: 'solid', color: '9B1B30', width: 2, dashType: 'dash' } } }
	),
	cellOn(
		'addTable.cell.diagonal.tlToBr.transparency',
		{ diagonal: { tlToBr: { type: 'solid', color: '9B1B30', width: 2 } } },
		{ diagonal: { tlToBr: { type: 'solid', color: '9B1B30', width: 2, transparency: 40 } } }
	),
	cell('addTable.cell.diagonal.blToTr', { diagonal: { blToTr: { type: 'solid', color: '9B1B30', width: 2 } } }, [
		'addTable.cell.diagonal.blToTr.type',
		'addTable.cell.diagonal.blToTr.color',
		'addTable.cell.diagonal.blToTr.width',
	]),
	cellOn(
		'addTable.cell.diagonal.blToTr.cap:round',
		{ diagonal: { blToTr: { type: 'solid', color: '9B1B30', width: 2 } } },
		{ diagonal: { blToTr: { type: 'solid', color: '9B1B30', width: 2, cap: 'round' } } }
	),
	cellOn(
		'addTable.cell.diagonal.blToTr.dashType:dash',
		{ diagonal: { blToTr: { type: 'solid', color: '9B1B30', width: 2 } } },
		{ diagonal: { blToTr: { type: 'solid', color: '9B1B30', width: 2, dashType: 'dash' } } }
	),
	cellOn(
		'addTable.cell.diagonal.blToTr.transparency',
		{ diagonal: { blToTr: { type: 'solid', color: '9B1B30', width: 2 } } },
		{ diagonal: { blToTr: { type: 'solid', color: '9B1B30', width: 2, transparency: 40 } } }
	),
	cell('addTable.cell.fill.type:solid', { fill: { type: 'solid', color: 'F0C808' } }, [
		'addTable.cell.fill',
		'addTable.cell.fill.type',
	]),
	cellOn('addTable.cell.fill.type:none', { fill: { color: 'F0C808' } }, { fill: { type: 'none' } }),
	cellOn(
		'addTable.cell.fill.transparency',
		{ fill: { color: 'F0C808' } },
		{ fill: { color: 'F0C808', transparency: 40 } }
	),
	cell(
		'addTable.cell.fill.gradient.linear',
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
			'addTable.cell.fill.gradient.kind',
			'addTable.cell.fill.gradient.angle',
			'addTable.cell.fill.gradient.stops.color',
			'addTable.cell.fill.gradient.stops.position',
			'addTable.cell.fill.gradient.stops.transparency',
		]
	),
	cellOn(
		'addTable.cell.fill.gradient.scaled',
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
	cellOn(
		'addTable.cell.fill.gradient.rotateWithShape',
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
	cell(
		'addTable.cell.fill.gradient.center',
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
		['addTable.cell.fill.gradient.center.x', 'addTable.cell.fill.gradient.center.y']
	),
	cell(
		'addTable.cell.fill.pattern',
		{ fill: { pattern: { preset: 'dkDnDiag', fgColor: '250F6B', bgColor: 'F0C808' } } },
		['addTable.cell.fill.pattern.preset', 'addTable.cell.fill.pattern.fgColor', 'addTable.cell.fill.pattern.bgColor']
	),
	cell('addTable.cell.fill.image.data', { fill: { image: { data: PNG } } }),
	cellOn(
		'addTable.cell.fill.image.crop',
		{ fill: { image: { data: PNG } } },
		{ fill: { image: { data: PNG, crop: { l: 10, t: 20, r: 30, b: 5 } } } },
		[
			'addTable.cell.fill.image.crop.l',
			'addTable.cell.fill.image.crop.t',
			'addTable.cell.fill.image.crop.r',
			'addTable.cell.fill.image.crop.b',
		]
	),
	{
		...cell('addTable.cell.fit:shrink', { fit: 'shrink' }),
		unwritten: 'https://github.com/shbernal/pptx-ts/issues/84',
	},
	{
		...cell('addTable.cell.fit.shrink', { fit: { type: 'shrink', fontScale: 80, lnSpcReduction: 10 } }, [
			'addTable.cell.fit.type',
			'addTable.cell.fit.fontScale',
			'addTable.cell.fit.lnSpcReduction',
		]),
		unwritten: 'https://github.com/shbernal/pptx-ts/issues/84',
	},
	cell('addTable.cell.fontFace', { fontFace: 'Georgia' }),
	cellOn('addTable.cell.fontFaceEA', { fontFace: 'Georgia' }, { fontFaceEA: 'MS Mincho' }),
	cell('addTable.cell.fontSize', { fontSize: 28 }),
	cell('addTable.cell.highlight', { highlight: 'FFFF00' }),
	cell('addTable.cell.horzOverflow:overflow', { horzOverflow: 'overflow' }),
	cell('addTable.cell.hyperlink.url', { hyperlink: { url: 'https://example.com/', tooltip: 'Example' } }, [
		'addTable.cell.hyperlink.url',
		'addTable.cell.hyperlink.tooltip',
	]),
	cell('addTable.cell.hyperlink.slide', { hyperlink: { slide: 1 } }),
	cell('addTable.cell.hyperlink.action:nextslide', { hyperlink: { action: 'nextslide' } }),
	cell('addTable.cell.indentLevel', { indentLevel: 2 }),
	cell('addTable.cell.italic', { italic: true }),
	cell('addTable.cell.lang', { lang: 'fr-FR' }),
	cell('addTable.cell.lineSpacing', { lineSpacing: 30 }),
	cell('addTable.cell.lineSpacingMultiple', { lineSpacingMultiple: 1.5 }),
	cell('addTable.cell.margin', { margin: [10, 20, 10, 20] }),
	cell('addTable.cell.paraIndent', { paraIndent: 12 }),
	cell('addTable.cell.paraMarginLeft', { paraMarginLeft: 18 }),
	cell('addTable.cell.paraSpaceAfter', { paraSpaceAfter: 12 }),
	cell('addTable.cell.paraSpaceBefore', { paraSpaceBefore: 12 }),
	cellRun('addTable.cell.softBreakBefore', 1, { softBreakBefore: true }),
	cell('addTable.cell.strike:dblStrike', { strike: 'dblStrike' }),
	cell('addTable.cell.tabStops', { tabStops: [{ position: 1.5, alignment: 'r' }] }, [
		'addTable.cell.tabStops.position',
		'addTable.cell.tabStops.alignment',
	]),
	cell('addTable.cell.textDirection:vert270', { textDirection: 'vert270' }),
	{
		...cell('addTable.cell.textWarp', { textWarp: 'textArchUp' }),
		unwritten: 'https://github.com/shbernal/pptx-ts/issues/84',
	},
	cell('addTable.cell.transparency', { transparency: 40 }),
	cell('addTable.cell.underline.style:dbl', { underline: { style: 'dbl' } }),
	cellOn(
		'addTable.cell.underline.color',
		{ underline: { style: 'sng' } },
		{ underline: { style: 'sng', color: '9B1B30' } }
	),
	cell('addTable.cell.valign:bottom', { valign: 'bottom' }),
	slide(
		'slide.background.color',
		(target) => {
			target.background = { color: 'F0C808' }
		},
		['slide.background', 'slide.background.color']
	),
	slide('slide.background:shorthand', (target) => {
		target.background = '250F6B'
	}),
	slideOn(
		'slide.background.transparency',
		(target) => {
			target.background = { color: 'F0C808' }
		},
		(target) => {
			target.background = { color: 'F0C808', transparency: 40 }
		}
	),
	slideOn(
		'slide.background.type:none',
		(target) => {
			target.background = { color: 'F0C808' }
		},
		(target) => {
			target.background = { color: 'F0C808', type: 'none' }
		},
		['slide.background.type']
	),
	slide(
		'slide.background.gradient.linear',
		(target) => {
			target.background = {
				gradient: {
					kind: 'linear',
					angle: 45,
					stops: [
						{ position: 0, color: 'F0C808' },
						{ position: 100, color: '250F6B', transparency: 30 },
					],
				},
			}
		},
		[
			'slide.background.gradient.kind',
			'slide.background.gradient.angle',
			'slide.background.gradient.stops.color',
			'slide.background.gradient.stops.position',
			'slide.background.gradient.stops.transparency',
		]
	),
	slideOn(
		'slide.background.gradient.scaled',
		(target) => {
			target.background = {
				gradient: {
					kind: 'linear',
					angle: 45,
					stops: [
						{ position: 0, color: 'F0C808' },
						{ position: 100, color: '250F6B' },
					],
				},
			}
		},
		(target) => {
			target.background = {
				gradient: {
					kind: 'linear',
					angle: 45,
					scaled: true,
					stops: [
						{ position: 0, color: 'F0C808' },
						{ position: 100, color: '250F6B' },
					],
				},
			}
		}
	),
	slideOn(
		'slide.background.gradient.rotateWithShape',
		(target) => {
			target.background = {
				gradient: {
					kind: 'linear',
					angle: 45,
					stops: [
						{ position: 0, color: 'F0C808' },
						{ position: 100, color: '250F6B' },
					],
				},
			}
		},
		(target) => {
			target.background = {
				gradient: {
					kind: 'linear',
					angle: 45,
					rotateWithShape: false,
					stops: [
						{ position: 0, color: 'F0C808' },
						{ position: 100, color: '250F6B' },
					],
				},
			}
		}
	),
	slide(
		'slide.background.gradient.center',
		(target) => {
			target.background = {
				gradient: {
					kind: 'radial',
					center: { x: 20, y: 70 },
					stops: [
						{ position: 0, color: 'F0C808' },
						{ position: 100, color: '250F6B' },
					],
				},
			}
		},
		['slide.background.gradient.center.x', 'slide.background.gradient.center.y']
	),
	slide(
		'slide.background.pattern',
		(target) => {
			target.background = { pattern: { preset: 'dkDnDiag', fgColor: '250F6B', bgColor: 'F0C808' } }
		},
		['slide.background.pattern.preset', 'slide.background.pattern.fgColor', 'slide.background.pattern.bgColor']
	),
	slide('slide.background.data', (target) => {
		target.background = { data: PNG }
	}),
	slide('slide.background.image.data', (target) => {
		target.background = { image: { data: PNG } }
	}),
	slideOn(
		'slide.background.image.crop',
		(target) => {
			target.background = { image: { data: PNG } }
		},
		(target) => {
			target.background = { image: { data: PNG, crop: { l: 10, t: 20, r: 30, b: 5 } } }
		},
		[
			'slide.background.image.crop.l',
			'slide.background.image.crop.t',
			'slide.background.image.crop.r',
			'slide.background.image.crop.b',
		]
	),
	slide('slide.color', (target) => {
		target.color = '9B1B30'
	}),
	slide('slide.hidden', (target) => {
		target.hidden = true
	}),
	slide('slide.transition.type', (target) => {
		target.transition = { type: 'push' }
	}),
	slideOn(
		'slide.transition.durationMs',
		(target) => {
			target.transition = { type: 'fade' }
		},
		(target) => {
			target.transition = { type: 'fade', durationMs: 1250 }
		}
	),
	slideOn(
		'slide.transition.speed:slow',
		(target) => {
			target.transition = { type: 'fade' }
		},
		(target) => {
			target.transition = { type: 'fade', speed: 'slow' }
		},
		['slide.transition.speed']
	),
	slideOn(
		'slide.transition.advanceOnClick',
		(target) => {
			target.transition = { type: 'fade' }
		},
		(target) => {
			target.transition = { type: 'fade', advanceOnClick: false }
		}
	),
	slideOn(
		'slide.transition.advanceAfterMs',
		(target) => {
			target.transition = { type: 'fade' }
		},
		(target) => {
			target.transition = { type: 'fade', advanceAfterMs: 3000 }
		}
	),
	slideOn(
		'slide.transition.variant',
		(target) => {
			target.transition = { type: 'push' }
		},
		(target) => {
			target.transition = { type: 'push', variant: { dir: 'u' } }
		}
	),
	slideOn(
		'slide.transition.sound.data',
		(target) => {
			target.transition = { type: 'fade' }
		},
		(target) => {
			target.transition = { type: 'fade', sound: { data: WAV } }
		}
	),
	slideOn(
		'slide.transition.sound.name',
		(target) => {
			target.transition = { type: 'fade', sound: { data: WAV } }
		},
		(target) => {
			target.transition = { type: 'fade', sound: { data: WAV, name: 'chime.wav' } }
		}
	),
	slideOn(
		'slide.transition.sound.loop',
		(target) => {
			target.transition = { type: 'fade', sound: { data: WAV } }
		},
		(target) => {
			target.transition = { type: 'fade', sound: { data: WAV, loop: true } }
		}
	),
	slideOn(
		'slide.transition.sound.stopPrevious',
		(target) => {
			target.transition = { type: 'fade' }
		},
		(target) => {
			target.transition = { type: 'fade', sound: { stopPrevious: true } }
		}
	),
	slideNumber('slide.slideNumber.x', { x: 2 }),
	slideNumber('slide.slideNumber.y', { y: 5 }),
	slideNumber('slide.slideNumber.w', { w: 2 }),
	slideNumber('slide.slideNumber.h', { h: 0.8 }),
	slideNumber('slide.slideNumber.align:center', { align: 'center' }, {}, ['slide.slideNumber.align']),
	slideNumber('slide.slideNumber.valign:bottom', { valign: 'bottom' }, {}, ['slide.slideNumber.valign']),
	slideNumber('slide.slideNumber.bold', { bold: true }),
	slideNumber('slide.slideNumber.italic', { italic: true }),
	slideNumber('slide.slideNumber.caps:all', { caps: 'all' }, {}, ['slide.slideNumber.caps']),
	slideNumber('slide.slideNumber.color', { color: '9B1B30' }),
	slideNumber('slide.slideNumber.fontFace', { fontFace: 'Georgia' }),
	slideNumber('slide.slideNumber.fontFaceEA', { fontFaceEA: 'MS Mincho' }),
	slideNumber('slide.slideNumber.fontSize', { fontSize: 20 }),
	slideNumber('slide.slideNumber.highlight', { highlight: 'F0C808' }),
	slideNumber('slide.slideNumber.lang', { lang: 'fr-FR' }),
	slideNumber('slide.slideNumber.margin', { margin: [0.1, 0.2, 0.1, 0.2] }),
	slideNumber('slide.slideNumber.transparency', { transparency: 40 }, { color: '9B1B30' }),
	slideNumber('slide.slideNumber.underline', { underline: { style: 'dbl', color: '9B1B30' } }, {}, [
		'slide.slideNumber.underline.style',
		'slide.slideNumber.underline.color',
	]),
	slideNumber('slide.slideNumber.textDirection:vert', { textDirection: 'vert' }, {}, [
		'slide.slideNumber.textDirection',
	]),
	slideNumber('slide.slideNumber.textWarp', { textWarp: 'textArchUp' }),
	slideNumber('slide.slideNumber.paraIndent', { paraIndent: 12 }),
	slideNumber('slide.slideNumber.paraMarginLeft', { paraMarginLeft: 18 }),
	{
		id: 'addNotes.text',
		covers: ['addNotes.text'],
		author: (pptx: TsPptx, on: boolean) => {
			const slide = pptx.addSlide()
			if (on) slide.addNotes('Probe notes')
		},
	},
	notes('addNotes.options.bold', { bold: true }),
	notes('addNotes.options.italic', { italic: true }),
	notes('addNotes.options.color', { color: '9B1B30' }),
	notes('addNotes.options.fontSize', { fontSize: 20 }),
	notes('addNotes.options.fontFace', { fontFace: 'Georgia' }),
	notes('addNotes.options.underline', { underline: { style: 'dbl', color: '9B1B30' } }, [
		'addNotes.options.underline.style',
		'addNotes.options.underline.color',
	]),
	notes('addNotes.options.hyperlink.url', { hyperlink: { url: 'https://example.com/', tooltip: 'Example' } }, [
		'addNotes.options.hyperlink.url',
		'addNotes.options.hyperlink.tooltip',
	]),
	{
		...notes('addNotes.options.hyperlink.slide', { hyperlink: { slide: 1 } }),
		unwritten: 'https://github.com/shbernal/pptx-ts/issues/168',
	},
	{
		...notes('addNotes.options.hyperlink.action', { hyperlink: { action: 'nextslide' } }),
		unwritten: 'https://github.com/shbernal/pptx-ts/issues/168',
	},
	master('defineSlideMaster.background.color', { background: { color: 'F0C808' } }, {}, [
		'defineSlideMaster.background',
		'defineSlideMaster.background.color',
	]),
	master(
		'defineSlideMaster.background.transparency',
		{ background: { color: 'F0C808', transparency: 40 } },
		{ background: { color: 'F0C808' } }
	),
	master(
		'defineSlideMaster.background.type:none',
		{ background: { color: 'F0C808', type: 'none' } },
		{ background: { color: 'F0C808' } },
		['defineSlideMaster.background.type']
	),
	master(
		'defineSlideMaster.background.gradient.linear',
		{
			background: {
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
		{},
		[
			'defineSlideMaster.background.gradient.kind',
			'defineSlideMaster.background.gradient.angle',
			'defineSlideMaster.background.gradient.stops.color',
			'defineSlideMaster.background.gradient.stops.position',
			'defineSlideMaster.background.gradient.stops.transparency',
		]
	),
	master(
		'defineSlideMaster.background.gradient.scaled',
		{ background: { gradient: { ...LINEAR, scaled: true } } },
		{ background: { gradient: LINEAR } }
	),
	master(
		'defineSlideMaster.background.gradient.rotateWithShape',
		{ background: { gradient: { ...LINEAR, rotateWithShape: false } } },
		{ background: { gradient: LINEAR } }
	),
	master(
		'defineSlideMaster.background.gradient.center',
		{
			background: {
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
		{},
		['defineSlideMaster.background.gradient.center.x', 'defineSlideMaster.background.gradient.center.y']
	),
	master(
		'defineSlideMaster.background.pattern',
		{ background: { pattern: { preset: 'dkDnDiag', fgColor: '250F6B', bgColor: 'F0C808' } } },
		{},
		[
			'defineSlideMaster.background.pattern.preset',
			'defineSlideMaster.background.pattern.fgColor',
			'defineSlideMaster.background.pattern.bgColor',
		]
	),
	master('defineSlideMaster.background.data', { background: { data: PNG } }),
	master('defineSlideMaster.background.image.data', { background: { image: { data: PNG } } }),
	master(
		'defineSlideMaster.background.image.crop',
		{ background: { image: { data: PNG, crop: { l: 10, t: 20, r: 30, b: 5 } } } },
		{ background: { image: { data: PNG } } },
		[
			'defineSlideMaster.background.image.crop.l',
			'defineSlideMaster.background.image.crop.t',
			'defineSlideMaster.background.image.crop.r',
			'defineSlideMaster.background.image.crop.b',
		]
	),
	{
		id: 'defineSlideMaster.margin',
		covers: ['defineSlideMaster.margin'],
		// The master margin writes nothing of its own: it is the deck-wide inset an
		// auto-paged table sizes itself within, so the probe places one on the master.
		author: (pptx: TsPptx, on: boolean) => {
			pptx.defineSlideMaster({ title: 'PROBE', ...(on ? { margin: [0.5, 1.5, 0.5, 1.5] } : {}) })
			pptx.addSlide({ masterTitle: 'PROBE' }).addTable(
				[
					[{ text: 'a' }, { text: 'b' }],
					[{ text: 'c' }, { text: 'd' }],
				],
				{ x: 1, y: 1, autoPage: true, objectName: 'probe' }
			)
		},
	},
	master('defineSlideMaster.objects.rect', { objects: [{ rect: { ...BOX, fill: { color: 'F0C808' } } }] }),
	master('defineSlideMaster.objects.roundRect', {
		objects: [{ roundRect: { ...BOX, fill: { color: 'F0C808' }, rectRadius: 0.2 } }],
	}),
	master('defineSlideMaster.objects.line', {
		objects: [{ line: { x: 1, y: 3, w: 5, h: 0, line: { color: '9B1B30', width: 2 } } }],
	}),
	master('defineSlideMaster.objects.shape', {
		objects: [{ shape: { type: 'ellipse', options: { ...BOX, fill: { color: 'F0C808' } } } }],
	}),
	master('defineSlideMaster.objects.text', {
		objects: [{ text: { text: 'Master text', options: { x: 1, y: 3, w: 5, h: 1 } } }],
	}),
	master('defineSlideMaster.objects.image', { objects: [{ image: { x: 1, y: 3, w: 1, h: 1, data: PNG } }] }),
	{
		...master('defineSlideMaster.objects.placeholder.options', {
			objects: [{ placeholder: { options: { name: 'body', type: 'body', x: 1, y: 3, w: 5, h: 1 } } }],
		}),
		issue: 'https://github.com/shbernal/pptx-ts/issues/165',
	},
	master(
		'defineSlideMaster.objects.placeholder.text',
		{
			objects: [
				{ placeholder: { options: { name: 'body', type: 'body', x: 1, y: 3, w: 5, h: 1 }, text: 'Body here' } },
			],
		},
		{ objects: [{ placeholder: { options: { name: 'body', type: 'body', x: 1, y: 3, w: 5, h: 1 } } }] }
	),
	master(
		'defineSlideMaster.objects.chart',
		{ objects: [{ chart: { type: 'bar', data: [SERIES], options: CHART_BOX } }] },
		{},
		[
			'defineSlideMaster.objects.chart.data.name',
			'defineSlideMaster.objects.chart.data.labels',
			'defineSlideMaster.objects.chart.data.values',
			'defineSlideMaster.objects.chart.options',
		]
	),
	masterChart('defineSlideMaster.objects.chart.data.customLabels', { customLabels: ['Low', '', 'High'] }),
	masterChart(
		'defineSlideMaster.objects.chart.data.pointStyles',
		{
			pointStyles: [{ fill: '00B050', border: { width: 2, color: '9B1B30' } }, {}, { pattern: { preset: 'ltUpDiag' } }],
		},
		[
			'defineSlideMaster.objects.chart.data.pointStyles.fill',
			'defineSlideMaster.objects.chart.data.pointStyles.border',
			'defineSlideMaster.objects.chart.data.pointStyles.pattern',
		]
	),
	masterChart(
		'defineSlideMaster.objects.chart.data.errorBars',
		{
			errorBars: {
				valueType: 'fixedVal',
				value: 2,
				barType: 'plus',
				direction: 'y',
				noEndCap: true,
				color: '9B1B30',
				width: 2,
				dashType: 'dash',
				cap: 'round',
				transparency: 30,
			},
		},
		[
			'defineSlideMaster.objects.chart.data.errorBars.valueType',
			'defineSlideMaster.objects.chart.data.errorBars.value',
			'defineSlideMaster.objects.chart.data.errorBars.barType',
			'defineSlideMaster.objects.chart.data.errorBars.direction',
			'defineSlideMaster.objects.chart.data.errorBars.noEndCap',
			'defineSlideMaster.objects.chart.data.errorBars.color',
			'defineSlideMaster.objects.chart.data.errorBars.width',
			'defineSlideMaster.objects.chart.data.errorBars.dashType',
			'defineSlideMaster.objects.chart.data.errorBars.cap',
			'defineSlideMaster.objects.chart.data.errorBars.transparency',
		]
	),
	masterChart(
		'defineSlideMaster.objects.chart.data.errorBars.cust',
		{ errorBars: { valueType: 'cust', plusValues: [1, 2, 1], minusValues: [0.5, 1, 0.5] } },
		[
			'defineSlideMaster.objects.chart.data.errorBars.plusValues',
			'defineSlideMaster.objects.chart.data.errorBars.minusValues',
		]
	),
	masterChart('defineSlideMaster.objects.chart.data.errorBars.size', { errorBars: { size: 3 } }),
	masterChart('defineSlideMaster.objects.chart.data.errorBars.type:dash', { errorBars: { type: 'dash' } }, [
		'defineSlideMaster.objects.chart.data.errorBars.type',
	]),
	master(
		'defineSlideMaster.objects.chart.data.sizes',
		{
			objects: [{ chart: { type: 'bubble', data: [BUBBLE_X, { ...BUBBLE_Y, sizes: [5, 1, 5] }], options: CHART_BOX } }],
		},
		{ objects: [{ chart: { type: 'bubble', data: [BUBBLE_X, BUBBLE_Y], options: CHART_BOX } }] }
	),
	master(
		'defineSlideMaster.objects.chart.type:combo',
		{ objects: [{ chart: { type: COMBO, data: [], options: COMBO_BOX } }] },
		{},
		[
			'defineSlideMaster.objects.chart.type.type',
			'defineSlideMaster.objects.chart.type.data.name',
			'defineSlideMaster.objects.chart.type.data.labels',
			'defineSlideMaster.objects.chart.type.data.values',
			'defineSlideMaster.objects.chart.type.options',
		]
	),
	master(
		'defineSlideMaster.objects.chart.type.data.series',
		{
			objects: [
				{
					chart: {
						type: [
							{
								type: 'bar',
								data: [
									{
										...SERIES,
										customLabels: ['Low', '', 'High'],
										pointStyles: [{ fill: '00B050' }],
										errorBars: { valueType: 'percentage', value: 5 },
									},
								],
								options: {},
							},
							...COMBO.slice(1),
						],
						data: [],
						options: COMBO_BOX,
					},
				},
			],
		},
		{ objects: [{ chart: { type: COMBO, data: [], options: COMBO_BOX } }] },
		[
			'defineSlideMaster.objects.chart.type.data.customLabels',
			'defineSlideMaster.objects.chart.type.data.pointStyles',
			'defineSlideMaster.objects.chart.type.data.errorBars',
		]
	),
	{
		id: 'defineSlideMaster.title',
		covers: ['defineSlideMaster.title'],
		author: (pptx: TsPptx, on: boolean) => {
			const title = on ? 'Probe master' : 'PROBE'
			pptx.defineSlideMaster({ title })
			pptx.addSlide({ masterTitle: title }).addText('Probe text', { ...BOX, objectName: 'probe' })
		},
	},
	masterNumber('defineSlideMaster.slideNumber.x', { x: 2 }),
	masterNumber('defineSlideMaster.slideNumber.y', { y: 5 }),
	masterNumber('defineSlideMaster.slideNumber.w', { w: 2 }),
	masterNumber('defineSlideMaster.slideNumber.h', { h: 0.8 }),
	masterNumber('defineSlideMaster.slideNumber.align:center', { align: 'center' }, {}, [
		'defineSlideMaster.slideNumber.align',
	]),
	masterNumber('defineSlideMaster.slideNumber.valign:bottom', { valign: 'bottom' }, {}, [
		'defineSlideMaster.slideNumber.valign',
	]),
	masterNumber('defineSlideMaster.slideNumber.bold', { bold: true }),
	masterNumber('defineSlideMaster.slideNumber.italic', { italic: true }),
	masterNumber('defineSlideMaster.slideNumber.caps:all', { caps: 'all' }, {}, ['defineSlideMaster.slideNumber.caps']),
	masterNumber('defineSlideMaster.slideNumber.color', { color: '9B1B30' }),
	masterNumber('defineSlideMaster.slideNumber.fontFace', { fontFace: 'Georgia' }),
	masterNumber('defineSlideMaster.slideNumber.fontFaceEA', { fontFaceEA: 'MS Mincho' }),
	masterNumber('defineSlideMaster.slideNumber.fontSize', { fontSize: 20 }),
	masterNumber('defineSlideMaster.slideNumber.highlight', { highlight: 'F0C808' }),
	masterNumber('defineSlideMaster.slideNumber.lang', { lang: 'fr-FR' }),
	masterNumber('defineSlideMaster.slideNumber.margin', { margin: [0.1, 0.2, 0.1, 0.2] }),
	masterNumber('defineSlideMaster.slideNumber.transparency', { transparency: 40 }, { color: '9B1B30' }),
	masterNumber('defineSlideMaster.slideNumber.underline', { underline: { style: 'dbl', color: '9B1B30' } }, {}, [
		'defineSlideMaster.slideNumber.underline.style',
		'defineSlideMaster.slideNumber.underline.color',
	]),
	masterNumber('defineSlideMaster.slideNumber.textDirection:vert', { textDirection: 'vert' }, {}, [
		'defineSlideMaster.slideNumber.textDirection',
	]),
	masterNumber('defineSlideMaster.slideNumber.textWarp', { textWarp: 'textArchUp' }),
	masterNumber('defineSlideMaster.slideNumber.paraIndent', { paraIndent: 12 }),
	masterNumber('defineSlideMaster.slideNumber.paraMarginLeft', { paraMarginLeft: 18 }),
	masterStyle('defineSlideMaster.textStyles.title.align:right', 'title', { align: 'right' }, {}, [
		'defineSlideMaster.textStyles.title.align',
	]),
	masterStyle('defineSlideMaster.textStyles.title.bold', 'title', { bold: true }),
	masterStyle('defineSlideMaster.textStyles.title.italic', 'title', { italic: true }),
	masterStyle('defineSlideMaster.textStyles.title.color', 'title', { color: '9B1B30' }),
	masterStyle('defineSlideMaster.textStyles.title.fontFace', 'title', { fontFace: 'Georgia' }),
	masterStyle('defineSlideMaster.textStyles.title.fontSize', 'title', { fontSize: 28 }),
	masterStyle('defineSlideMaster.textStyles.title.marginLeft', 'title', { marginLeft: 0.5 }),
	masterStyle('defineSlideMaster.textStyles.title.indent', 'title', { indent: -0.25 }, { marginLeft: 0.5 }),
	masterStyle('defineSlideMaster.textStyles.title.bullet:true', 'title', { bullet: true }, {}, [
		'defineSlideMaster.textStyles.title.bullet',
	]),
	masterStyle('defineSlideMaster.textStyles.title.bullet.characterCode', 'title', {
		bullet: { characterCode: '25BA' },
	}),
	masterStyle(
		'defineSlideMaster.textStyles.title.bullet.fontFace',
		'title',
		{ bullet: { characterCode: '25BA', fontFace: 'Wingdings' } },
		{ bullet: { characterCode: '25BA' } }
	),
	masterStyle('defineSlideMaster.textStyles.title.bullet.type:number', 'title', { bullet: { type: 'number' } }, {}, [
		'defineSlideMaster.textStyles.title.bullet.type',
	]),
	masterStyle(
		'defineSlideMaster.textStyles.title.bullet.numberType',
		'title',
		{ bullet: { type: 'number', numberType: 'romanLcParenR' } },
		{ bullet: { type: 'number' } }
	),
	masterStyle(
		'defineSlideMaster.textStyles.title.bullet.numberStartAt',
		'title',
		{ bullet: { type: 'number', numberStartAt: 5 } },
		{ bullet: { type: 'number' } }
	),
	masterStyle('defineSlideMaster.textStyles.body.align:center', 'body', { align: 'center' }, {}, [
		'defineSlideMaster.textStyles.body.align',
	]),
	masterStyle('defineSlideMaster.textStyles.body.bold', 'body', { bold: true }),
	masterStyle('defineSlideMaster.textStyles.body.italic', 'body', { italic: true }),
	masterStyle('defineSlideMaster.textStyles.body.color', 'body', { color: '9B1B30' }),
	masterStyle('defineSlideMaster.textStyles.body.fontFace', 'body', { fontFace: 'Georgia' }),
	masterStyle('defineSlideMaster.textStyles.body.fontSize', 'body', { fontSize: 28 }),
	masterStyle('defineSlideMaster.textStyles.body.marginLeft', 'body', { marginLeft: 0.5 }),
	masterStyle('defineSlideMaster.textStyles.body.indent', 'body', { indent: -0.25 }, { marginLeft: 0.5 }),
	masterStyle('defineSlideMaster.textStyles.body.bullet:false', 'body', { bullet: false }, {}, [
		'defineSlideMaster.textStyles.body.bullet',
	]),
	masterStyle('defineSlideMaster.textStyles.body.bullet.characterCode', 'body', { bullet: { characterCode: '25BA' } }),
	masterStyle(
		'defineSlideMaster.textStyles.body.bullet.fontFace',
		'body',
		{ bullet: { characterCode: '25BA', fontFace: 'Wingdings' } },
		{ bullet: { characterCode: '25BA' } }
	),
	masterStyle('defineSlideMaster.textStyles.body.bullet.type:number', 'body', { bullet: { type: 'number' } }, {}, [
		'defineSlideMaster.textStyles.body.bullet.type',
	]),
	masterStyle(
		'defineSlideMaster.textStyles.body.bullet.numberType',
		'body',
		{ bullet: { type: 'number', numberType: 'romanLcParenR' } },
		{ bullet: { type: 'number' } }
	),
	masterStyle(
		'defineSlideMaster.textStyles.body.bullet.numberStartAt',
		'body',
		{ bullet: { type: 'number', numberStartAt: 5 } },
		{ bullet: { type: 'number' } }
	),
	masterStyle('defineSlideMaster.textStyles.other.align:center', 'other', { align: 'center' }, {}, [
		'defineSlideMaster.textStyles.other.align',
	]),
	masterStyle('defineSlideMaster.textStyles.other.bold', 'other', { bold: true }),
	masterStyle('defineSlideMaster.textStyles.other.italic', 'other', { italic: true }),
	masterStyle('defineSlideMaster.textStyles.other.color', 'other', { color: '9B1B30' }),
	masterStyle('defineSlideMaster.textStyles.other.fontFace', 'other', { fontFace: 'Georgia' }),
	masterStyle('defineSlideMaster.textStyles.other.fontSize', 'other', { fontSize: 28 }),
	masterStyle('defineSlideMaster.textStyles.other.marginLeft', 'other', { marginLeft: 0.5 }),
	masterStyle('defineSlideMaster.textStyles.other.indent', 'other', { indent: -0.25 }, { marginLeft: 0.5 }),
	masterStyle('defineSlideMaster.textStyles.other.bullet:false', 'other', { bullet: false }, {}, [
		'defineSlideMaster.textStyles.other.bullet',
	]),
	masterStyle('defineSlideMaster.textStyles.other.bullet.characterCode', 'other', {
		bullet: { characterCode: '25BA' },
	}),
	masterStyle(
		'defineSlideMaster.textStyles.other.bullet.fontFace',
		'other',
		{ bullet: { characterCode: '25BA', fontFace: 'Wingdings' } },
		{ bullet: { characterCode: '25BA' } }
	),
	masterStyle('defineSlideMaster.textStyles.other.bullet.type:number', 'other', { bullet: { type: 'number' } }, {}, [
		'defineSlideMaster.textStyles.other.bullet.type',
	]),
	masterStyle(
		'defineSlideMaster.textStyles.other.bullet.numberType',
		'other',
		{ bullet: { type: 'number', numberType: 'romanLcParenR' } },
		{ bullet: { type: 'number' } }
	),
	masterStyle(
		'defineSlideMaster.textStyles.other.bullet.numberStartAt',
		'other',
		{ bullet: { type: 'number', numberStartAt: 5 } },
		{ bullet: { type: 'number' } }
	),
	image('addImage.altText', { altText: 'A probe description' }),
	image('addImage.biLevel.threshold', { biLevel: { threshold: 0.3 } }),
	image('addImage.clrChange', { clrChange: { from: '000000', to: '9B1B30' } }, [
		'addImage.clrChange.from',
		'addImage.clrChange.to',
	]),
	image('addImage.crop', { crop: { l: 10, t: 20, r: 30, b: 15 } }, [
		'addImage.crop.l',
		'addImage.crop.t',
		'addImage.crop.r',
		'addImage.crop.b',
	]),
	image('addImage.duotone', { duotone: { shadow: '250F6B', highlight: 'F0C808' } }, [
		'addImage.duotone.shadow',
		'addImage.duotone.highlight',
	]),
	image('addImage.flipH', { flipH: true }),
	image('addImage.flipV', { flipV: true }),
	image('addImage.grayscale', { grayscale: true }),
	image('addImage.x', { x: 2 }),
	image('addImage.y', { y: 2 }),
	image('addImage.w', { w: 6 }),
	image('addImage.h', { h: 2 }),
	image('addImage.hyperlink.url', { hyperlink: { url: 'https://example.com/', tooltip: 'Example' } }, [
		'addImage.hyperlink.url',
		'addImage.hyperlink.tooltip',
	]),
	image('addImage.hyperlink.slide', { hyperlink: { slide: 1 } }),
	image('addImage.hyperlink.action:nextslide', { hyperlink: { action: 'nextslide' } }, ['addImage.hyperlink.action']),
	image('addImage.line', { line: { color: '250F6B', width: 2 } }, ['addImage.line.color', 'addImage.line.width']),
	imageOn(
		'addImage.line.dashType:dash',
		{ line: { color: '250F6B', width: 2 } },
		{ line: { color: '250F6B', width: 2, dashType: 'dash' } },
		['addImage.line.dashType']
	),
	imageOn(
		'addImage.line.cap:round',
		{ line: { color: '250F6B', width: 2 } },
		{ line: { color: '250F6B', width: 2, cap: 'round' } },
		['addImage.line.cap']
	),
	imageOn(
		'addImage.line.beginArrowType',
		{ line: { color: '250F6B', width: 2 } },
		{ line: { color: '250F6B', width: 2, beginArrowType: 'triangle' } }
	),
	imageOn(
		'addImage.line.endArrowType',
		{ line: { color: '250F6B', width: 2 } },
		{ line: { color: '250F6B', width: 2, endArrowType: 'oval' } }
	),
	imageOn(
		'addImage.line.transparency',
		{ line: { color: '250F6B', width: 2 } },
		{ line: { color: '250F6B', width: 2, transparency: 40 } }
	),
	imageOn(
		'addImage.line.type:none',
		{ line: { color: '250F6B', width: 2 } },
		{ line: { color: '250F6B', width: 2, type: 'none' } },
		['addImage.line.type']
	),
	image(
		'addImage.line.gradient.linear',
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
			'addImage.line.gradient.kind',
			'addImage.line.gradient.angle',
			'addImage.line.gradient.stops.color',
			'addImage.line.gradient.stops.position',
			'addImage.line.gradient.stops.transparency',
		]
	),
	imageOn(
		'addImage.line.gradient.scaled',
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
	imageOn(
		'addImage.line.gradient.rotateWithShape',
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
	image(
		'addImage.line.gradient.center',
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
		['addImage.line.gradient.center.x', 'addImage.line.gradient.center.y']
	),
	image(
		'addImage.line.pattern',
		{ line: { width: 3, pattern: { preset: 'dkDnDiag', fgColor: '250F6B', bgColor: 'F0C808' } } },
		['addImage.line.pattern.preset', 'addImage.line.pattern.fgColor', 'addImage.line.pattern.bgColor']
	),
	image('addImage.objectLock.noAdjustHandles', { objectLock: { noAdjustHandles: true } }),
	image('addImage.objectLock.noChangeArrowheads', { objectLock: { noChangeArrowheads: true } }),
	image('addImage.objectLock.noChangeAspect:false', { objectLock: { noChangeAspect: false } }, [
		'addImage.objectLock.noChangeAspect',
	]),
	image('addImage.objectLock.noChangeShapeType', { objectLock: { noChangeShapeType: true } }),
	image('addImage.objectLock.noCrop', { objectLock: { noCrop: true } }),
	{
		...image('addImage.objectLock.noDrilldown', { objectLock: { noDrilldown: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'the writer warns that noDrilldown is not supported on a:picLocks and drops it: it is a graphic-frame/group lock, which a picture does not have',
		},
	},
	image('addImage.objectLock.noEditPoints', { objectLock: { noEditPoints: true } }),
	image('addImage.objectLock.noGrp', { objectLock: { noGrp: true } }),
	image('addImage.objectLock.noMove', { objectLock: { noMove: true } }),
	image('addImage.objectLock.noResize', { objectLock: { noResize: true } }),
	image('addImage.objectLock.noRot', { objectLock: { noRot: true } }),
	image('addImage.objectLock.noSelect', { objectLock: { noSelect: true } }),
	{
		...image('addImage.objectLock.noTextEdit', { objectLock: { noTextEdit: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'the writer warns that noTextEdit is not supported on a:picLocks and drops it: a picture has no text body to lock',
		},
	},
	image('addImage.objectName', { objectName: 'Renamed probe' }),
	{
		id: 'addImage.placeholder',
		covers: ['addImage.placeholder'],
		author: (pptx: TsPptx, on: boolean) => {
			pptx.defineSlideMaster({
				title: 'PROBE_PICTURE_MASTER',
				objects: [{ placeholder: { options: { name: 'pic', type: 'pic', x: 1, y: 1, w: 5, h: 1.5 }, text: '' } }],
			})
			pptx
				.addSlide({ masterTitle: 'PROBE_PICTURE_MASTER' })
				.addImage({ data: PNG, ...BOX, objectName: 'probe', ...(on ? { placeholder: 'pic' } : {}) })
		},
	},
	image('addImage.points', { points: [{ x: 0, y: 0 }, { x: 5, y: 0 }, { x: 2.5, y: 1.5 }, { close: true }] }, [
		'addImage.points.x',
		'addImage.points.y',
		'addImage.points.close',
	]),
	imageOn(
		'addImage.points.moveTo',
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
	imageOn(
		'addImage.points.curve',
		{
			points: [
				{ x: 0, y: 0 },
				{ x: 5, y: 0 },
			],
		},
		{ points: [{ x: 0, y: 0 }, { x: 5, y: 0 }, { curve: { type: 'arc', hR: 0.75, wR: 2.5, stAng: 0, swAng: 180 } }] },
		[
			'addImage.points.curve.type',
			'addImage.points.curve.hR',
			'addImage.points.curve.wR',
			'addImage.points.curve.stAng',
			'addImage.points.curve.swAng',
		]
	),
	imageOn('addImage.rectRadius', { shape: 'roundRect' }, { rectRadius: 0.3 }),
	image('addImage.rotate', { rotate: 30 }),
	image('addImage.rounding', { rounding: true }),
	image(
		'addImage.shadow.outer',
		{ shadow: { type: 'outer', angle: 60, blur: 6, color: '250F6B', offset: 4, transparency: 40 } },
		[
			'addImage.shadow.type',
			'addImage.shadow.angle',
			'addImage.shadow.blur',
			'addImage.shadow.color',
			'addImage.shadow.offset',
			'addImage.shadow.transparency',
		]
	),
	imageOn(
		'addImage.shadow.rotateWithShape',
		{ shadow: { type: 'outer', blur: 6 } },
		{ shadow: { type: 'outer', blur: 6, rotateWithShape: true } }
	),
	image('addImage.shape:hexagon', { shape: 'hexagon' }, ['addImage.shape']),
	imageOn('addImage.shapeAdjust', { shape: 'roundRect' }, { shapeAdjust: { name: 'adj', value: 0.4 } }, [
		'addImage.shapeAdjust.name',
		'addImage.shapeAdjust.value',
	]),
	image('addImage.sizing.type:cover', { sizing: { type: 'cover' } }, ['addImage.sizing.type']),
	image('addImage.sizing.contain.wh', { sizing: { type: 'contain', w: 2, h: 1 } }, [
		'addImage.sizing.w',
		'addImage.sizing.h',
	]),
	image('addImage.sizing.crop.xy', { sizing: { type: 'crop', x: 1, y: 0.5, w: 2, h: 0.5 } }, [
		'addImage.sizing.x',
		'addImage.sizing.y',
	]),
	{
		id: 'addImage.svg',
		covers: ['addImage.svg'],
		author: (pptx: TsPptx, on: boolean) => {
			// A flat SVG of the probe box's aspect, against the PNG raster on the off side.
			const svg =
				'<svg xmlns="http://www.w3.org/2000/svg" width="10" height="3" viewBox="0 0 10 3"><rect width="10" height="3" fill="#9B1B30"/></svg>'
			pptx.addSlide().addImage(on ? { svg, ...BOX, objectName: 'probe' } : { data: PNG, ...BOX, objectName: 'probe' })
		},
	},
	image('addImage.transparency', { transparency: 40 }),
	connector('addConnector.type:elbow', {}, { type: 'elbow' }, ['addConnector.type']),
	connector('addConnector.bends', { type: 'elbow' }, { bends: 2 }),
	connector('addConnector.adj', { type: 'elbow' }, { adj: 30 }),
	connector('addConnector.color', {}, { color: '9B1B30' }),
	connector('addConnector.width', {}, { width: 3 }),
	connector('addConnector.dashType:dash', {}, { dashType: 'dash' }, ['addConnector.dashType']),
	connector('addConnector.beginArrowType', {}, { beginArrowType: 'triangle' }),
	connector('addConnector.endArrowType', {}, { endArrowType: 'triangle' }),
	connector('addConnector.startShape', {}, { startShape: 'from' }),
	connector('addConnector.startShapeIdx', { startShape: 'from' }, { startShapeIdx: 2 }),
	connector('addConnector.endShape', {}, { endShape: 'to' }),
	connector('addConnector.endShapeIdx', { endShape: 'to' }, { endShapeIdx: 2 }),
	connector('addConnector.objectName', {}, { objectName: 'Renamed probe' }),
	connector('addConnector.altText', {}, { altText: 'A connector between two boxes' }),
	connector('addConnector.x1', {}, { x1: 2 }),
	connector('addConnector.y1', {}, { y1: 2 }),
	connector('addConnector.x2', {}, { x2: 6 }),
	connector('addConnector.y2', {}, { y2: 4 }),
	group('addGroup.x', GROUP_FRAME, { x: 0.5 }),
	group('addGroup.y', GROUP_FRAME, { y: 0.5 }),
	group('addGroup.w', GROUP_FRAME, { w: 5 }),
	group('addGroup.h', GROUP_FRAME, { h: 2 }),
	group('addGroup.rotate', {}, { rotate: 30 }),
	group('addGroup.flipH', {}, { flipH: true }),
	group('addGroup.flipV', {}, { flipV: true }),
	group('addGroup.objectName', {}, { objectName: 'Renamed probe' }),
	group('addGroup.altText', {}, { altText: 'Two boxes, grouped' }),
	{
		...group('addGroup.objectLock.noAdjustHandles', {}, { objectLock: { noAdjustHandles: true } }),
		accepted: NOT_A_GROUP_LOCK,
	},
	{
		...group('addGroup.objectLock.noChangeArrowheads', {}, { objectLock: { noChangeArrowheads: true } }),
		accepted: NOT_A_GROUP_LOCK,
	},
	group('addGroup.objectLock.noChangeAspect', {}, { objectLock: { noChangeAspect: true } }),
	{
		...group('addGroup.objectLock.noChangeShapeType', {}, { objectLock: { noChangeShapeType: true } }),
		accepted: NOT_A_GROUP_LOCK,
	},
	{ ...group('addGroup.objectLock.noCrop', {}, { objectLock: { noCrop: true } }), accepted: NOT_A_GROUP_LOCK },
	{
		...group('addGroup.objectLock.noDrilldown', {}, { objectLock: { noDrilldown: true } }),
		accepted: NOT_A_GROUP_LOCK,
	},
	{
		...group('addGroup.objectLock.noEditPoints', {}, { objectLock: { noEditPoints: true } }),
		accepted: NOT_A_GROUP_LOCK,
	},
	group('addGroup.objectLock.noGrp', {}, { objectLock: { noGrp: true } }),
	group('addGroup.objectLock.noMove', {}, { objectLock: { noMove: true } }),
	group('addGroup.objectLock.noResize', {}, { objectLock: { noResize: true } }),
	group('addGroup.objectLock.noRot', {}, { objectLock: { noRot: true } }),
	group('addGroup.objectLock.noSelect', {}, { objectLock: { noSelect: true } }),
	{ ...group('addGroup.objectLock.noTextEdit', {}, { objectLock: { noTextEdit: true } }), accepted: NOT_A_GROUP_LOCK },
	theme('theme.headFontFace', { headFontFace: 'Georgia' }),
	theme('theme.bodyFontFace', { bodyFontFace: 'Verdana' }),
	theme('theme.headFontFaceEA', { headFontFaceEA: 'Yu Gothic' }),
	theme('theme.bodyFontFaceEA', { bodyFontFaceEA: 'MS Mincho' }),
	theme('theme.headFontFaceCS', { headFontFaceCS: 'Tahoma' }),
	theme('theme.bodyFontFaceCS', { bodyFontFaceCS: 'Arial' }),
	theme('theme.colorScheme.dk1', { colorScheme: { dk1: '101820' } }),
	theme('theme.colorScheme.lt1', { colorScheme: { lt1: 'FAF7F0' } }),
	theme('theme.colorScheme.dk2', { colorScheme: { dk2: '250F6B' } }),
	theme('theme.colorScheme.lt2', { colorScheme: { lt2: 'DDE3F0' } }),
	theme('theme.colorScheme.accent1', { colorScheme: { accent1: '9B1B30' } }),
	theme('theme.colorScheme.accent2', { colorScheme: { accent2: 'F0C808' } }),
	theme('theme.colorScheme.accent3', { colorScheme: { accent3: '2E8B57' } }),
	theme('theme.colorScheme.accent4', { colorScheme: { accent4: '8A2BE2' } }),
	theme('theme.colorScheme.accent5', { colorScheme: { accent5: 'FF7F50' } }),
	theme('theme.colorScheme.accent6', { colorScheme: { accent6: '20B2AA' } }),
	theme('theme.colorScheme.hlink', { colorScheme: { hlink: '0050A0' } }),
	theme('theme.colorScheme.folHlink', { colorScheme: { folHlink: '7A3E9D' } }),
]

/**
 * Inventory paths deliberately not probed, and why. A path here counts as covered,
 * so the reason has to hold up: "hard to probe" is not one; "selects the input
 * source and writes nothing of its own" is.
 */
export const SKIPPED: Record<string, string> = {
	'addImage.data':
		'the picture source itself: every addImage probe writes its picture from it, so it has no option-free side to compare against',
	'addImage.path':
		'the file-path spelling of the picture source; the picture it writes is exercised by every addImage probe through data',
	'addText.data': 'an image source for picture-filled text, exercised by addText.fill.image.*',
	'addText.path': 'an image source for picture-filled text, exercised by addText.fill.image.*',
	'addText.fill.image.path':
		'the file-path spelling of the picture source; the picture it writes is exercised by addText.fill.image.data',
	'addText.bullet.image.path':
		'the file-path spelling of the bullet picture source; the bullet it writes is exercised by addText.bullet.image.data',
	'addShape.fill.image.path':
		'the file-path spelling of the picture source; the picture it writes is exercised by addShape.fill.image.data',
	'addTable.cell.fill.image.path':
		'the file-path spelling of the picture source; the picture it writes is exercised by addTable.cell.fill.image.data',
	'addTable.cell.bullet.image.path':
		'the file-path spelling of the bullet picture source; the bullet it writes is exercised by addTable.cell.bullet.image.data',
	'addTable.cell.autoPageCharWeight':
		'an auto-paging knob that only weighs where autoPage splits a table, and writes no XML of its own',
	'slide.background.path':
		'the file-path spelling of the background picture source; the picture it writes is exercised by slide.background.data',
	'slide.background.image.path':
		'the file-path spelling of the picture source; the picture it writes is exercised by slide.background.image.data',
	'slide.transition.sound.path':
		'the file-path spelling of the sound source; the sound it writes is exercised by slide.transition.sound.data',
	'defineSlideMaster.background.path':
		'the file-path spelling of the background picture source; the picture it writes is exercised by defineSlideMaster.background.data',
	'defineSlideMaster.background.image.path':
		'the file-path spelling of the picture source; the picture it writes is exercised by defineSlideMaster.background.image.data',
	'defineSlideMaster.objects.chart.type.data.sizes':
		'bubble sizes are read only by a bubble chart, and the writer refuses a bubble chart in the multi-type form by design (InvalidOptionError: PowerPoint refuses the deck), so this spelling cannot be authored (ts-pptx#140: https://github.com/shbernal/pptx-ts/issues/140 asks for the type to stop admitting it); the sizes it would write are exercised by defineSlideMaster.objects.chart.data.sizes',
}
