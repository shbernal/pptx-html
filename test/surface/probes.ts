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
	AnimationProps,
	CommentProps,
	CHART_NAME,
	ChartOpts,
	ConnectorProps,
	GroupProps,
	ImageProps,
	MediaProps,
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

/** A text box on a slide whose master has a `body` placeholder; `placeholder` binds it on the probe side, over `base`. */
function textPlaceholder(
	id: string,
	base: TextPropsOptions['placeholder'],
	placeholder: TextPropsOptions['placeholder']
): Probe {
	return {
		id,
		covers: [id.split(':')[0] as string],
		author: (pptx: TsPptx, on: boolean) => {
			pptx.defineSlideMaster({
				title: 'PROBE_BODY_MASTER',
				objects: [{ placeholder: { options: { name: 'body', type: 'body', ...BOX }, text: '' } }],
			})
			const bound = on ? placeholder : base
			const options: TextPropsOptions = { ...BOX, objectName: 'probe' }
			if (bound !== undefined) options.placeholder = bound
			pptx.addSlide({ masterTitle: 'PROBE_BODY_MASTER' }).addText('Probe text', options)
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

/** A 2×2 table with `base` on the table call on both sides, for options that only mean something on top of another. */
function tableOn(id: string, base: TableProps, options: TableProps, covers = [id.split(':')[0] as string]): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addTable(
				[
					[{ text: 'a' }, { text: 'b' }],
					[{ text: 'c' }, { text: 'd' }],
				],
				{ ...BOX, objectName: 'probe', ...base, ...(on ? options : {}) }
			)
		},
	}
}

/** A 2×2 table on a master with a table placeholder; the table binds to `base` off and `placeholder` on. */
function tablePlaceholder(id: string, base: TableProps['placeholder'], placeholder: TableProps['placeholder']): Probe {
	return {
		id,
		covers: [id.split(':')[0] as string],
		author: (pptx: TsPptx, on: boolean) => {
			pptx.defineSlideMaster({
				title: 'PROBE_TABLE_MASTER',
				objects: [{ placeholder: { options: { name: 'tbl', type: 'tbl', ...BOX }, text: '' } }],
			})
			const bound = on ? placeholder : base
			const options: TableProps = { ...BOX, objectName: 'probe' }
			if (bound !== undefined) options.placeholder = bound
			pptx.addSlide({ masterTitle: 'PROBE_TABLE_MASTER' }).addTable(
				[
					[{ text: 'a' }, { text: 'b' }],
					[{ text: 'c' }, { text: 'd' }],
				],
				options
			)
		},
	}
}

/**
 * A table of twenty tall rows that overflows one slide, with `base` on both sides,
 * for the auto-paging options.
 */
function tallTable(id: string, base: TableProps, options: TableProps, covers = [id.split(':')[0] as string]): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addTable(
				Array.from({ length: 20 }, (_, row) => [{ text: `r${row}a` }, { text: `r${row}b` }]),
				{ x: 1, y: 1, w: 5, rowH: 0.5, objectName: 'probe', ...base, ...(on ? options : {}) }
			)
		},
	}
}

/**
 * A 1.8 KB TrueType face, "Probe Mono", built with opentype.js: lowercase letters and a
 * space, each 600 units wide with no outline. Measured fit only runs once a face has
 * metrics registered, so a fit probe registers this one.
 */
const FONT =
	'T1RUTwAKAIAAAwAgQ0ZGIFBavrcAAAUoAAABYU9TLzJmlWAMAAABEAAAAGBjbWFwAM0AHAAABNQAAAA0aGVhZC9VJWYAAACsAAAANmhoZWEDIgGtAAAA5AAAACRobXR4QaAAAAAABowAAABwbHRhZ2V+AAQAAAb8AAAAEm1heHAAHFAAAAABCAAAAAZuYW1l3j3nbwAAAXAAAANjcG9zdAADAAAAAAUIAAAAIAABAAAAAQAAmwdrVF8PPPUAAwPoAAAAAObzckMAAAAA5vNyQwAAAAACWAAAAAAAAwACAAAAAAAAAAEAAAMg/zgAAAJYAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAcAABQAAAcAAAAAwJYAfQABQAAAooCuwAAAIwCigK7AAAB3wAxAQIAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAWFhYWABAACAAegMg/zgAAAAAAAAAAAABAAAAAAAAAAAAIAAgAAAAAAAzAmoAAAAEAAAAAAACAAoAAAAEAAAAAQAUAAAAAAAEAAAAAgAOABQAAAAEAAAAAwAqAHwAAAAEAAAABAAkACIAAAAEAAAABQAWAGYAAAAEAAAABgAgAEYAAAAEAAAABwACAAoAAAAEAAAACAACAAoAAAAEAAAACQACAAoAAAAEAAAACgACAAoAAAAEAAAACwACAAoAAAAEAAAADAACAAoAAAAEAAAADQACAAoAAAAEAAAADgACAAoAAAAEAAAAEAAUAAAAAAAEAAAAEQAOABQAAQAAAAAAAAABAAsAAQAAAAAAAQAKAKYAAQAAAAAAAgAHALAAAQAAAAAAAwAVAOQAAQAAAAAABAASALcAAQAAAAAABQALANkAAQAAAAAABgAQAMkAAQAAAAAABwABAAsAAQAAAAAACAABAAsAAQAAAAAACQABAAsAAQAAAAAACgABAAsAAQAAAAAACwABAAsAAQAAAAAADAABAAsAAQAAAAAADQABAAsAAQAAAAAADgABAAsAAQAAAAAAEAAKAKYAAQAAAAAAEQAHALAAAwABBAkAAAACAAoAAwABBAkAAQAUAAAAAwABBAkAAgAOABQAAwABBAkAAwAqAHwAAwABBAkABAAkACIAAwABBAkABQAWAGYAAwABBAkABgAgAEYAAwABBAkABwACAAoAAwABBAkACAACAAoAAwABBAkACQACAAoAAwABBAkACgACAAoAAwABBAkACwACAAoAAwABBAkADAACAAoAAwABBAkADQACAAoAAwABBAkADgACAAoAAwABBAkAEAAUAAAAAwABBAkAEQAOABQAUAByAG8AYgBlACAATQBvAG4AbwBSAGUAZwB1AGwAYQByAFAAcgBvAGIAZQAgAE0AbwBuAG8AIABSAGUAZwB1AGwAYQByAFAAcgBvAGIAZQBNAG8AbgBvAFIAZQBnAHUAbABhAHIAVgBlAHIAcwBpAG8AbgAgADAALgAxACAAOgAgAFAAcgBvAGIAZQAgAE0AbwBuAG8AIABSAGUAZwB1AGwAYQByUHJvYmUgTW9ub1JlZ3VsYXJQcm9iZSBNb25vIFJlZ3VsYXJQcm9iZU1vbm9SZWd1bGFyVmVyc2lvbiAwLjEgOiBQcm9iZSBNb25vIFJlZ3VsYXIAAAAAAQADAAEAAAAMAAQAKAAAAAYABAABAAIAIAB6//8AAAAgAGH////h/6EAAQAAAAAAAAADAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAEAQABAQERUHJvYmVNb25vUmVndWxhcgABAQEn+BsA+BwC+B0D+B4Ei4v5tPjsBR0AAAC2Dx0AAADtEYsdAAABYRIAHwEBDB4oLzQ1Njc4OTo7PD0+P0BBQkNERUZHSElKS0xNTlZlcnNpb24gMC4xUHJvYmUgTW9ubyBSZWd1bGFyUHJvYmUgTW9ub1JlZ3VsYXJzcGFjZWFiY2RlZmdoaWprbG1ub3BxcnN0dXZ3eHl6AAAAAYsBjAGNAY4BjwGQAZEBkgGTAZQBlQGWAZcBmAGZAZoBmwGcAZ0BngGfAaABoQGiAaMBpAGlABwBAQQHCg0QExYZHB8iJSgrLjE0Nzo9QENGSUxPUlX47A747A747A747A747A747A747A747A747A747A747A747A747A747A747A747A747A747A747A747A747A747A747A747A747A747A747A747A4AAAACWAAAAlgAAAJYAAACWAAAAlgAAAJYAAACWAAAAlgAAAJYAAACWAAAAlgAAAJYAAACWAAAAlgAAAJYAAACWAAAAlgAAAJYAAACWAAAAlgAAAJYAAACWAAAAlgAAAJYAAACWAAAAlgAAAJYAAACWAAAAAAAAQAAAAAAAAABABAAAmVuAAA='

/** A 2×2 table in 60 pt Probe Mono, whose metrics are registered on both sides, so a measured fit has text to shrink. */
function fitTable(id: string, options: TableProps, covers = [id.split(':')[0] as string]): Probe {
	return {
		id,
		covers,
		author: async (pptx: TsPptx, on: boolean) => {
			await pptx.registerFontMetrics(
				'Probe Mono',
				Uint8Array.from(atob(FONT), (char) => char.charCodeAt(0))
			)
			pptx.addSlide().addTable(
				[
					[{ text: 'a' }, { text: 'b' }],
					[{ text: 'c' }, { text: 'd' }],
				],
				{ ...BOX, objectName: 'probe', fontFace: 'Probe Mono', fontSize: 60, ...(on ? options : {}) }
			)
		},
	}
}

/** {@link fitTable} with `options` on the first cell instead of the table. */
function fitCell(id: string, options: TableCellProps, covers = [id.split(':')[0] as string]): Probe {
	return {
		id,
		covers,
		author: async (pptx: TsPptx, on: boolean) => {
			await pptx.registerFontMetrics(
				'Probe Mono',
				Uint8Array.from(atob(FONT), (char) => char.charCodeAt(0))
			)
			pptx.addSlide().addTable(
				[
					[{ text: 'a', options: on ? options : {} }, { text: 'b' }],
					[{ text: 'c' }, { text: 'd' }],
				],
				{ ...BOX, objectName: 'probe', fontFace: 'Probe Mono', fontSize: 60 }
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

/** Two series over three categories: enough for axes, legends and per-series options to mean something. */
const CHART_SERIES: OptsChartData[] = [
	{ name: 'North', labels: ['Q1', 'Q2', 'Q3'], values: [4, 7, 5] },
	{ name: 'South', labels: ['Q1', 'Q2', 'Q3'], values: [3, 5, 8] },
]

/** One series with a negative point, which is what `invertedColors` colours. */
const NEGATIVE: OptsChartData[] = [{ name: 'Net', labels: ['Q1', 'Q2', 'Q3'], values: [4, -2, 5] }]

/** An X row, then one series with its bubble sizes. */
const BUBBLES: OptsChartData[] = [
	{ name: 'X', values: [1, 2, 3] },
	{ name: 'North', values: [4, 7, 5], sizes: [3, 6, 4] },
]

/** `addChart` of `type` over `data`, with `base` on both sides; `options` are merged in on the probe side only. */
function chartWith(
	id: string,
	type: CHART_NAME,
	data: OptsChartData[],
	base: ChartOpts,
	options: ChartOpts,
	covers = [id.split(':')[0] as string]
): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addChart(data, { ...BOX, h: 4, objectName: 'probe', ...base, ...(on ? options : {}), type })
		},
	}
}

/** `addChart` of `type` over {@link CHART_SERIES}, with `base` on both sides; `options` are merged in on the probe side only. */
function chartOn(
	id: string,
	type: CHART_NAME,
	base: ChartOpts,
	options: ChartOpts,
	covers = [id.split(':')[0] as string]
): Probe {
	return chartWith(id, type, CHART_SERIES, base, options, covers)
}

/** A bar chart; `options` are applied only when the probe is on. */
function chart(id: string, options: ChartOpts, covers = [id.split(':')[0] as string]): Probe {
	return chartOn(id, 'bar', {}, options, covers)
}

/** ISO-dated categories, which a `catLabelFormatCode` turns into a date axis. */
const DATED: OptsChartData[] = [
	{ name: 'North', labels: ['2026-01-01', '2026-02-01', '2026-03-01'], values: [4, 7, 5] },
	{ name: 'South', labels: ['2026-01-01', '2026-02-01', '2026-03-01'], values: [3, 5, 8] },
]

/** Two category levels, for the multi-level label options. */
const LEVELS: OptsChartData[] = [
	{
		name: 'North',
		labels: [
			['Q1', 'Q2', 'Q3', 'Q4'],
			['H1', '', 'H2', ''],
		],
		values: [4, 7, 5, 6],
	},
]

/** High, low and close, in the order `stockStyle: 'hlc'` reads them. */
const STOCK: OptsChartData[] = [
	{ name: 'High', labels: ['Mon', 'Tue', 'Wed'], values: [12, 14, 13] },
	{ name: 'Low', labels: ['Mon', 'Tue', 'Wed'], values: [8, 9, 10] },
	{ name: 'Close', labels: ['Mon', 'Tue', 'Wed'], values: [10, 12, 11] },
]

/** Raw observations, which a histogram or a box-and-whisker chart summarizes itself. */
const OBSERVATIONS: OptsChartData[] = [{ name: 'Sample', values: [1, 2, 2, 3, 3, 3, 4, 4, 5, 9, 15] }]

/** Region names, for a region map. */
const REGIONS: OptsChartData[] = [{ name: 'Sales', labels: ['France', 'Germany', 'Spain'], values: [4, 7, 5] }]

/** The line half of a bar + line combo. */
const TREND: OptsChartData[] = [{ name: 'Trend', labels: ['Q1', 'Q2', 'Q3'], values: [5, 6, 7] }]

/** A bar + line combo; `options` go on the line subchart only when the probe is on. */
function chartCombo(id: string, options: ChartOpts, covers = [id.split(':')[0] as string]): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addChart(
				[
					{ type: 'bar', data: CHART_SERIES, options: {} },
					{ type: 'line', data: TREND, options: on ? options : {} },
				],
				{ ...BOX, h: 4, objectName: 'probe' }
			)
		},
	}
}

/** A bar chart on a slide whose master has a `chart` placeholder; `placeholder` binds it on the probe side, over `base`. */
function chartPlaceholder(id: string, base: ChartOpts['placeholder'], placeholder: ChartOpts['placeholder']): Probe {
	return {
		id,
		covers: [id.split(':')[0] as string],
		author: (pptx: TsPptx, on: boolean) => {
			pptx.defineSlideMaster({
				title: 'PROBE_CHART_MASTER',
				objects: [{ placeholder: { options: { name: 'chart', type: 'chart', x: 1, y: 1, w: 5, h: 4 }, text: '' } }],
			})
			const bound = on ? placeholder : base
			const options: ChartOpts = { ...BOX, h: 4, objectName: 'probe' }
			if (bound !== undefined) options.placeholder = bound
			pptx.addSlide({ masterTitle: 'PROBE_CHART_MASTER' }).addChart(CHART_SERIES, { ...options, type: 'bar' })
		},
	}
}

/** A silent eight-sample WAV, for transition sounds. */
const WAV = 'audio/wav;base64,UklGRiwAAABXQVZFZm10IBAAAAABAAEAQB8AAEAfAAABAAgAZGF0YQgAAACAgICAgICAgA=='

const MEDIA_BOX = { x: 1, y: 1, w: 4, h: 3 }

/** An embedded sound with `base` on both sides and `options` merged in on the probe side. */
function media(
	id: string,
	options: Partial<MediaProps>,
	base: Partial<MediaProps> = {},
	covers = [id.split(':')[0] as string]
): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addMedia({
				type: 'audio',
				data: WAV,
				...MEDIA_BOX,
				objectName: 'probe',
				...base,
				...(on ? options : {}),
			} as MediaProps)
		},
	}
}

const COMMENT: CommentProps = { author: 'Ada Probe', text: 'Probe comment' }

/**
 * A slide holding the probe text box and one review comment, with `options` merged into the
 * comment on the probe side. With no `options` the comment itself is what is probed: the off
 * side has none.
 */
function comment(id: string, options: Partial<CommentProps>, covers = [id.split(':')[0] as string]): Probe {
	const bare = Object.keys(options).length === 0
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			const target = pptx.addSlide()
			target.addText('Probe text', { ...BOX, objectName: 'probe' })
			if (on || !bare) target.addComment({ ...COMMENT, ...(on ? options : {}) })
		},
	}
}

/**
 * A slide holding two text boxes, the probe and a second target, with `base` animating one of
 * them on both sides and `options` merged in on the probe side. With no `base` the animation
 * itself is what is probed: the off side has none.
 */
function animation(
	id: string,
	options: Partial<AnimationProps>,
	base?: AnimationProps,
	covers = [id.split(':')[0] as string]
): Probe {
	return {
		id,
		covers,
		author: (pptx: TsPptx, on: boolean) => {
			const target = pptx.addSlide()
			target.addText('Probe text', { ...BOX, objectName: 'probe' })
			target.addText('Second text', { ...BOX, y: 3, objectName: 'second' })
			if (on || base) target.addAnimation({ ...base, ...(on ? options : {}) } as AnimationProps)
		},
	}
}

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

/** The PNG on a slide whose master has a `pic` placeholder; `placeholder` binds it on the probe side, over `base`. */
function imagePlaceholder(id: string, base: ImageProps['placeholder'], placeholder: ImageProps['placeholder']): Probe {
	return {
		id,
		covers: [id.split(':')[0] as string],
		author: (pptx: TsPptx, on: boolean) => {
			pptx.defineSlideMaster({
				title: 'PROBE_PICTURE_MASTER',
				objects: [{ placeholder: { options: { name: 'pic', type: 'pic', ...BOX }, text: '' } }],
			})
			const bound = on ? placeholder : base
			const options: ImageProps = { data: PNG, ...BOX, objectName: 'probe' }
			if (bound !== undefined) options.placeholder = bound
			pptx.addSlide({ masterTitle: 'PROBE_PICTURE_MASTER' }).addImage(options)
		},
	}
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

const AUTHORED: Probe[] = [
	text('addText.bold', { bold: true }),
	text('addText.italic', { italic: true }),
	text('addText.underline.style:dbl', { underline: { style: 'dbl' } }),
	text('addText.underline:true', { underline: true }),
	text('addText.strike:dblStrike', { strike: 'dblStrike' }),
	text('addText.charSpacing', { charSpacing: 4 }),
	text('addText.kerning', { kerning: 20 }),
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
	{
		// The probe above has only Latin text, which the East Asian face never
		// paints; this one gives it characters to paint.
		id: 'addText.fontFaceEA:cjk',
		covers: ['addText.fontFaceEA'],
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addText('プローブ', {
				...BOX,
				objectName: 'probe',
				fontFace: 'Georgia',
				...(on ? { fontFaceEA: 'MS Mincho' } : {}),
			})
		},
	},
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
	{
		...text('addText.isTextBox', { isTextBox: true }),
		invisible:
			'`p:cNvSpPr/@txBox`: marks the shape as a text box, which changes how PowerPoint edits and autofits it on the next change, and draws nothing',
	},
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
	{
		...text('addText.objectLock.noUngrp', { objectLock: { noUngrp: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'the writer drops noUngrp on a:spLocks: it is a group lock (a:grpSpLocks only), which a text box does not have',
		},
	},
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
	{
		// The probe above has no tab character, so its stops have nothing to move;
		// this one gives them one.
		id: 'addText.tabStops:tabbed',
		covers: ['addText.tabStops.position', 'addText.tabStops.alignment'],
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addText('Probe\ttext', {
				...BOX,
				objectName: 'probe',
				...(on ? { tabStops: [{ position: 1.5, alignment: 'r' as const }] } : {}),
			})
		},
	},
	text('addText.textDirection:vert', { textDirection: 'vert' }),
	text('addText.textWarp', { textWarp: 'textArchUp' }),
	text('addText.transparency', { transparency: 40 }),
	textOn('addText.underline.color', { underline: { style: 'sng' } }, { underline: { style: 'sng', color: '9B1B30' } }),
	text('addText.vert:eaVert', { vert: 'eaVert' }),
	text('addText.wrap', { wrap: false }),
	text('addText.placeholder', { placeholder: 'body' }),
	textPlaceholder('addText.placeholder.name', undefined, { name: 'body' }),
	textPlaceholder('addText.placeholder.type', undefined, { type: 'body' }),
	textPlaceholder('addText.placeholder.idx', { type: 'body' }, { type: 'body', idx: 13 }),
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
	// The probe above is upright, where a gradient that does not turn with its
	// shape looks the same as one that does; this one turns the shape.
	shapeOn(
		'addShape.fill.gradient.rotateWithShape:rotated',
		'rect',
		{
			rotate: 30,
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
			rotate: 30,
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
	{
		...shape('addShape.objectLock.noUngrp', { objectLock: { noUngrp: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'the writer drops noUngrp on a:spLocks: it is a group lock (a:grpSpLocks only), which a shape does not have',
		},
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
		// A guide paints only through a path point that names it, which pptx-ts can
		// neither write nor read back: the reader turns `x="g0"` into 0. Once it can,
		// the probe that proves rendering is a new one whose path uses the guide.
		unread: 'https://github.com/shbernal/pptx-ts/issues/255',
	},
	{
		...shapeOn(
			'addShape.connectionSites',
			'custGeom',
			{ points: CUST },
			{ connectionSites: [{ ang: 90, x: 2.5, y: 1.5 }] },
			['addShape.connectionSites.ang', 'addShape.connectionSites.x', 'addShape.connectionSites.y']
		),
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
	cell('addTable.cell.charSpacing', { charSpacing: 4 }),
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
	fitCell('addTable.cell.fit:shrink', { fit: 'shrink' }),
	{
		...fitCell('addTable.cell.fit.shrink', { fit: { type: 'shrink', fontScale: 80, lnSpcReduction: 10 } }, [
			'addTable.cell.fit.type',
			'addTable.cell.fit.fontScale',
			'addTable.cell.fit.lnSpcReduction',
		]),
		accepted: {
			kind: 'inert',
			reason:
				'only fit: shrink is acted on in a cell (TableCellProps.fit): PowerPoint has no autofit in a cell, so the object form has no flag to write and a row already grows to fit its text (ts-pptx#84)',
		},
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
	cell('addTable.cell.kerning', { kerning: 20 }),
	cell('addTable.cell.lang', { lang: 'fr-FR' }),
	cell('addTable.cell.lineSpacing', { lineSpacing: 30 }),
	cell('addTable.cell.lineSpacingMultiple', { lineSpacingMultiple: 1.5 }),
	cell('addTable.cell.margin', { margin: [10, 20, 10, 20] }),
	cell('addTable.cell.paraIndent', { paraIndent: 12 }),
	cell('addTable.cell.paraMarginLeft', { paraMarginLeft: 18 }),
	cell('addTable.cell.paraSpaceAfter', { paraSpaceAfter: 12 }),
	cell('addTable.cell.paraSpaceBefore', { paraSpaceBefore: 12 }),
	cell(
		'addTable.cell.shadow.outer',
		{ shadow: { type: 'outer', angle: 60, blur: 6, color: '250F6B', offset: 4, transparency: 40 } },
		[
			'addTable.cell.shadow.type',
			'addTable.cell.shadow.angle',
			'addTable.cell.shadow.blur',
			'addTable.cell.shadow.color',
			'addTable.cell.shadow.offset',
			'addTable.cell.shadow.transparency',
		]
	),
	cellOn(
		'addTable.cell.shadow.rotateWithShape',
		{ shadow: { type: 'outer', blur: 6 } },
		{ shadow: { type: 'outer', blur: 6, rotateWithShape: true } }
	),
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
	cell('addTable.cell.underline:true', { underline: true }),
	cellOn(
		'addTable.cell.underline.color',
		{ underline: { style: 'sng' } },
		{ underline: { style: 'sng', color: '9B1B30' } }
	),
	cell('addTable.cell.valign:bottom', { valign: 'bottom' }),
	table('addTable.altText', { altText: 'A probe description' }),
	table('addTable.align:center', { align: 'center' }, ['addTable.align']),
	{
		...tallTable('addTable.autoPage', {}, { autoPage: true }),
		accepted: {
			kind: 'equivalent',
			reason:
				'the only facts lost are the empty notes slide pptx-ts writes for the continuation slide; the round trip writes no notes part for a slide without notes in either lane, and an absent notes slide shows the same empty notes page',
		},
	},
	tallTable('addTable.autoPageRepeatHeader', { autoPage: true }, { autoPageRepeatHeader: true }),
	tallTable('addTable.autoPageHeaderRows', { autoPage: true, autoPageRepeatHeader: true }, { autoPageHeaderRows: 2 }),
	tallTable('addTable.autoPageSlideStartY', { autoPage: true }, { autoPageSlideStartY: 2 }),
	{
		id: 'addTable.autoPagePlaceholder',
		covers: ['addTable.autoPagePlaceholder'],
		author: (pptx: TsPptx, on: boolean) => {
			pptx.defineSlideMaster({
				title: 'PROBE_TITLE_MASTER',
				objects: [
					{ placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 0.2, w: 9, h: 0.6 }, text: '' } },
				],
			})
			const slide = pptx.addSlide({ masterTitle: 'PROBE_TITLE_MASTER' })
			slide.addText('Probe title', { placeholder: 'title' })
			slide.addTable(
				Array.from({ length: 20 }, (_, row) => [{ text: `r${row}a` }, { text: `r${row}b` }]),
				{
					x: 1,
					y: 1,
					w: 5,
					rowH: 0.5,
					objectName: 'probe',
					autoPage: true,
					...(on ? { autoPagePlaceholder: true } : {}),
				}
			)
		},
	},
	table('addTable.bold', { bold: true }),
	table('addTable.border.type:dash', { border: { type: 'dash', color: '9B1B30', width: 2 } }, ['addTable.border.type']),
	tableOn(
		'addTable.border.width',
		{ border: { type: 'solid', color: '9B1B30' } },
		{ border: { type: 'solid', color: '9B1B30', width: 4 } }
	),
	tableOn(
		'addTable.border.cap:round',
		{ border: { type: 'solid', color: '9B1B30', width: 2 } },
		{ border: { type: 'solid', color: '9B1B30', width: 2, cap: 'round' } }
	),
	tableOn(
		'addTable.border.dashType:dash',
		{ border: { type: 'solid', color: '9B1B30', width: 2 } },
		{ border: { type: 'solid', color: '9B1B30', width: 2, dashType: 'dash' } }
	),
	tableOn(
		'addTable.border.transparency',
		{ border: { type: 'solid', color: '9B1B30', width: 2 } },
		{ border: { type: 'solid', color: '9B1B30', width: 2, transparency: 40 } }
	),
	{
		id: 'addTable.breakLine',
		covers: ['addTable.breakLine'],
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addTable(
				[
					[{ text: [{ text: 'First run' }, { text: 'Second run' }] }, { text: 'b' }],
					[{ text: 'c' }, { text: 'd' }],
				],
				{ ...BOX, objectName: 'probe', ...(on ? { breakLine: true } : {}) }
			)
		},
	},
	table('addTable.bullet', { bullet: true }),
	tableOn('addTable.bullet.characterCode', { bullet: true }, { bullet: { characterCode: '25BA' } }),
	tableOn('addTable.bullet.color', { bullet: true }, { bullet: { color: '9B1B30' } }),
	tableOn('addTable.bullet.fontFace', { bullet: true }, { bullet: { fontFace: 'Wingdings' } }),
	tableOn('addTable.bullet.image.data', { bullet: true }, { bullet: { image: { data: PNG } } }),
	tableOn('addTable.bullet.indent', { bullet: true }, { bullet: { indent: 10 } }),
	tableOn(
		'addTable.bullet.numberStartAt',
		{ bullet: { type: 'number' } },
		{ bullet: { type: 'number', numberType: 'romanUcPeriod', numberStartAt: 3 } },
		['addTable.bullet.numberStartAt', 'addTable.bullet.numberType']
	),
	tableOn('addTable.bullet.size', { bullet: true }, { bullet: { size: 80 } }),
	table('addTable.bullet.type:number', { bullet: { type: 'number' } }, ['addTable.bullet.type']),
	table('addTable.caps:all', { caps: 'all' }, ['addTable.caps']),
	table('addTable.color', { color: '9B1B30' }),
	table('addTable.columns', { columns: [{ fill: { color: 'F0C808' } }] }),
	table('addTable.colW', { colW: [2, 3] }),
	table('addTable.fill.color', { fill: { color: 'F0C808' } }),
	table('addTable.fill.type:solid', { fill: { type: 'solid', color: 'F0C808' } }, [
		'addTable.fill',
		'addTable.fill.type',
	]),
	tableOn('addTable.fill.transparency', { fill: { color: 'F0C808' } }, { fill: { color: 'F0C808', transparency: 40 } }),
	table(
		'addTable.fill.gradient.linear',
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
			'addTable.fill.gradient.kind',
			'addTable.fill.gradient.angle',
			'addTable.fill.gradient.stops.color',
			'addTable.fill.gradient.stops.position',
			'addTable.fill.gradient.stops.transparency',
		]
	),
	tableOn(
		'addTable.fill.gradient.scaled',
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
	tableOn(
		'addTable.fill.gradient.rotateWithShape',
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
	table(
		'addTable.fill.gradient.center',
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
		['addTable.fill.gradient.center.x', 'addTable.fill.gradient.center.y']
	),
	table('addTable.fill.pattern', { fill: { pattern: { preset: 'dkDnDiag', fgColor: '250F6B', bgColor: 'F0C808' } } }, [
		'addTable.fill.pattern.preset',
		'addTable.fill.pattern.fgColor',
		'addTable.fill.pattern.bgColor',
	]),
	table('addTable.fill.image.data', { fill: { image: { data: PNG } } }),
	tableOn(
		'addTable.fill.image.crop',
		{ fill: { image: { data: PNG } } },
		{ fill: { image: { data: PNG, crop: { l: 10, t: 20, r: 30, b: 5 } } } },
		[
			'addTable.fill.image.crop.l',
			'addTable.fill.image.crop.t',
			'addTable.fill.image.crop.r',
			'addTable.fill.image.crop.b',
		]
	),
	fitTable('addTable.fit:shrink', { fit: 'shrink' }, ['addTable.fit']),
	{
		...fitTable('addTable.fit.shrink', { fit: { type: 'shrink', fontScale: 80, lnSpcReduction: 10 } }, [
			'addTable.fit.type',
			'addTable.fit.fontScale',
			'addTable.fit.lnSpcReduction',
		]),
		accepted: {
			kind: 'inert',
			reason:
				'only fit: shrink is acted on in a table (TableProps.fit): PowerPoint has no autofit in a cell, so the object form has no flag to write and a row already grows to fit its text (ts-pptx#84)',
		},
	},
	tableOn('addTable.fitColumns', { x: 1, colW: [6, 6] }, { fitColumns: 'shrink' }),
	table('addTable.fontFace', { fontFace: 'Georgia' }),
	tableOn('addTable.fontFaceEA', { fontFace: 'Georgia' }, { fontFaceEA: 'MS Mincho' }),
	table('addTable.fontSize', { fontSize: 28 }),
	table('addTable.h', { h: 2 }),
	tableOn(
		'addTable.hasBandedColumns',
		{ tableStyle: '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}' },
		{ hasBandedColumns: true }
	),
	tableOn('addTable.hasBandedRows', { tableStyle: '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}' }, { hasBandedRows: true }),
	tableOn(
		'addTable.hasFirstColumn',
		{ tableStyle: '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}' },
		{ hasFirstColumn: true }
	),
	tableOn('addTable.hasFooter', { tableStyle: '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}' }, { hasFooter: true }),
	tableOn('addTable.hasHeader', { tableStyle: '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}' }, { hasHeader: true }),
	tableOn('addTable.hasLastColumn', { tableStyle: '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}' }, { hasLastColumn: true }),
	table('addTable.headerRow', { headerRow: { bold: true, fill: { color: 'F0C808' } } }),
	table('addTable.highlight', { highlight: 'FFFF00' }),
	table('addTable.italic', { italic: true }),
	table('addTable.lang', { lang: 'fr-FR' }),
	table('addTable.margin', { margin: [10, 20, 10, 20] }),
	{
		...table('addTable.objectLock.noAdjustHandles', { objectLock: { noAdjustHandles: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'the writer warns that noAdjustHandles is not supported on a:graphicFrameLocks and drops it: a table frame carries only noGrp, noDrilldown, noSelect, noChangeAspect, noMove and noResize',
		},
	},
	{
		...table('addTable.objectLock.noChangeArrowheads', { objectLock: { noChangeArrowheads: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'the writer warns that noChangeArrowheads is not supported on a:graphicFrameLocks and drops it: a table frame carries only noGrp, noDrilldown, noSelect, noChangeAspect, noMove and noResize',
		},
	},
	table('addTable.objectLock.noChangeAspect', { objectLock: { noChangeAspect: true } }),
	{
		...table('addTable.objectLock.noChangeShapeType', { objectLock: { noChangeShapeType: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'the writer warns that noChangeShapeType is not supported on a:graphicFrameLocks and drops it: a table frame carries only noGrp, noDrilldown, noSelect, noChangeAspect, noMove and noResize',
		},
	},
	{
		...table('addTable.objectLock.noCrop', { objectLock: { noCrop: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'the writer warns that noCrop is not supported on a:graphicFrameLocks and drops it: a table frame carries only noGrp, noDrilldown, noSelect, noChangeAspect, noMove and noResize',
		},
	},
	table('addTable.objectLock.noDrilldown', { objectLock: { noDrilldown: true } }),
	{
		...table('addTable.objectLock.noEditPoints', { objectLock: { noEditPoints: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'the writer warns that noEditPoints is not supported on a:graphicFrameLocks and drops it: a table frame carries only noGrp, noDrilldown, noSelect, noChangeAspect, noMove and noResize',
		},
	},
	table('addTable.objectLock.noGrp:false', { objectLock: { noGrp: false } }, ['addTable.objectLock.noGrp']),
	table('addTable.objectLock.noMove', { objectLock: { noMove: true } }),
	table('addTable.objectLock.noResize', { objectLock: { noResize: true } }),
	{
		...table('addTable.objectLock.noRot', { objectLock: { noRot: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'the writer warns that noRot is not supported on a:graphicFrameLocks and drops it: a table frame carries only noGrp, noDrilldown, noSelect, noChangeAspect, noMove and noResize',
		},
	},
	table('addTable.objectLock.noSelect', { objectLock: { noSelect: true } }),
	{
		...table('addTable.objectLock.noTextEdit', { objectLock: { noTextEdit: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'the writer warns that noTextEdit is not supported on a:graphicFrameLocks and drops it: a table frame carries only noGrp, noDrilldown, noSelect, noChangeAspect, noMove and noResize',
		},
	},
	{
		...table('addTable.objectLock.noUngrp', { objectLock: { noUngrp: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'the writer warns that noUngrp is not supported on a:graphicFrameLocks and drops it: a table frame carries only noGrp, noDrilldown, noSelect, noChangeAspect, noMove and noResize',
		},
	},
	table('addTable.objectName', { objectName: 'Renamed probe' }),
	table('addTable.outerBorder', { outerBorder: { type: 'solid', color: '9B1B30', width: 2 } }, [
		'addTable.outerBorder.type',
		'addTable.outerBorder.color',
		'addTable.outerBorder.width',
	]),
	tableOn(
		'addTable.outerBorder.cap:round',
		{ outerBorder: { type: 'solid', color: '9B1B30', width: 2 } },
		{ outerBorder: { type: 'solid', color: '9B1B30', width: 2, cap: 'round' } }
	),
	tableOn(
		'addTable.outerBorder.dashType:dash',
		{ outerBorder: { type: 'solid', color: '9B1B30', width: 2 } },
		{ outerBorder: { type: 'solid', color: '9B1B30', width: 2, dashType: 'dash' } }
	),
	tableOn(
		'addTable.outerBorder.transparency',
		{ outerBorder: { type: 'solid', color: '9B1B30', width: 2 } },
		{ outerBorder: { type: 'solid', color: '9B1B30', width: 2, transparency: 40 } }
	),
	table('addTable.paraIndent', { paraIndent: 12 }),
	table('addTable.paraMarginLeft', { paraMarginLeft: 18 }),
	tablePlaceholder('addTable.placeholder', undefined, 'tbl'),
	tablePlaceholder('addTable.placeholder.name', undefined, { name: 'tbl' }),
	tablePlaceholder('addTable.placeholder.type', undefined, { type: 'tbl' }),
	tablePlaceholder('addTable.placeholder.idx', { type: 'tbl' }, { type: 'tbl', idx: 13 }),
	table('addTable.rowH', { rowH: [0.5, 1] }),
	table('addTable.rtl', { rtl: true }),
	{
		id: 'addTable.slideMargin',
		covers: ['addTable.slideMargin'],
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addTable(
				[
					[{ text: 'a' }, { text: 'b' }],
					[{ text: 'c' }, { text: 'd' }],
				],
				{ x: 1, y: 1, objectName: 'probe', ...(on ? { slideMargin: [0.5, 2, 0.5, 1] } : {}) }
			)
		},
	},
	{
		id: 'addTable.softBreakBefore',
		covers: ['addTable.softBreakBefore'],
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addTable(
				[
					[{ text: [{ text: 'First run' }, { text: 'Second run' }] }, { text: 'b' }],
					[{ text: 'c' }, { text: 'd' }],
				],
				{ ...BOX, objectName: 'probe', ...(on ? { softBreakBefore: true } : {}) }
			)
		},
	},
	table('addTable.tableFill.color', { tableFill: { color: 'F0C808' } }),
	table('addTable.tableFill.type:solid', { tableFill: { type: 'solid', color: 'F0C808' } }, [
		'addTable.tableFill',
		'addTable.tableFill.type',
	]),
	tableOn(
		'addTable.tableFill.transparency',
		{ tableFill: { color: 'F0C808' } },
		{ tableFill: { color: 'F0C808', transparency: 40 } }
	),
	table(
		'addTable.tableFill.gradient.linear',
		{
			tableFill: {
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
			'addTable.tableFill.gradient.kind',
			'addTable.tableFill.gradient.angle',
			'addTable.tableFill.gradient.stops.color',
			'addTable.tableFill.gradient.stops.position',
			'addTable.tableFill.gradient.stops.transparency',
		]
	),
	tableOn(
		'addTable.tableFill.gradient.scaled',
		{
			tableFill: {
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
			tableFill: {
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
	tableOn(
		'addTable.tableFill.gradient.rotateWithShape',
		{
			tableFill: {
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
			tableFill: {
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
	table(
		'addTable.tableFill.gradient.center',
		{
			tableFill: {
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
		['addTable.tableFill.gradient.center.x', 'addTable.tableFill.gradient.center.y']
	),
	table(
		'addTable.tableFill.pattern',
		{ tableFill: { pattern: { preset: 'dkDnDiag', fgColor: '250F6B', bgColor: 'F0C808' } } },
		['addTable.tableFill.pattern.preset', 'addTable.tableFill.pattern.fgColor', 'addTable.tableFill.pattern.bgColor']
	),
	table('addTable.tableFill.image.data', { tableFill: { image: { data: PNG } } }),
	tableOn(
		'addTable.tableFill.image.crop',
		{ tableFill: { image: { data: PNG } } },
		{ tableFill: { image: { data: PNG, crop: { l: 10, t: 20, r: 30, b: 5 } } } },
		[
			'addTable.tableFill.image.crop.l',
			'addTable.tableFill.image.crop.t',
			'addTable.tableFill.image.crop.r',
			'addTable.tableFill.image.crop.b',
		]
	),
	table('addTable.tableStyle', { tableStyle: '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}' }),
	table('addTable.tabStops', { tabStops: [{ position: 1.5, alignment: 'r' }] }, [
		'addTable.tabStops.position',
		'addTable.tabStops.alignment',
	]),
	table('addTable.textDirection:vert270', { textDirection: 'vert270' }, ['addTable.textDirection']),
	{
		...table('addTable.textWarp', { textWarp: 'textArchUp' }),
		accepted: {
			kind: 'inert',
			reason:
				'the writer warns table/option-not-supported and drops it: a table cell has no a:bodyPr warp, and PowerPoint refuses WarpFormat on a cell (ts-pptx#84)',
		},
	},
	table('addTable.transparency', { transparency: 40 }),
	table('addTable.underline.style:dbl', { underline: { style: 'dbl' } }, ['addTable.underline.style']),
	table('addTable.underline:true', { underline: true }),
	tableOn(
		'addTable.underline.color',
		{ underline: { style: 'sng' } },
		{ underline: { style: 'sng', color: '9B1B30' } }
	),
	table('addTable.valign:bottom', { valign: 'bottom' }, ['addTable.valign']),
	table('addTable.w', { w: 6 }),
	table('addTable.x', { x: 2 }),
	table('addTable.y', { y: 2 }),
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
	slideNumber('slide.slideNumber.underline:true', { underline: true }),
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
	notes('addNotes.options.underline:true', { underline: true }),
	notes('addNotes.options.hyperlink.url', { hyperlink: { url: 'https://example.com/', tooltip: 'Example' } }, [
		'addNotes.options.hyperlink.url',
		'addNotes.options.hyperlink.tooltip',
	]),
	{
		...notes('addNotes.options.hyperlink.slide', { hyperlink: { slide: 1 } }),
	},
	{
		...notes('addNotes.options.hyperlink.action', { hyperlink: { action: 'nextslide' } }),
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
	master('defineSlideMaster.objects.placeholder.options', {
		objects: [{ placeholder: { options: { name: 'body', type: 'body', x: 1, y: 3, w: 5, h: 1 } } }],
	}),
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
	masterNumber('defineSlideMaster.slideNumber.underline:true', { underline: true }),
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
	{
		...image('addImage.objectLock.noUngrp', { objectLock: { noUngrp: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'the writer drops noUngrp on a:picLocks: it is a group lock (a:grpSpLocks only), which a picture does not have',
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
	imagePlaceholder('addImage.placeholder.name', undefined, { name: 'pic' }),
	imagePlaceholder('addImage.placeholder.type', undefined, { type: 'pic' }),
	imagePlaceholder('addImage.placeholder.idx', { type: 'pic' }, { type: 'pic', idx: 13 }),
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
	group('addGroup.objectLock.noUngrp', {}, { objectLock: { noUngrp: true } }),
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
	media('addMedia.altText', { altText: 'A probe description' }),
	media('addMedia.cover', { cover: PNG }),
	media('addMedia.extn', { extn: 'mp3' }),
	media('addMedia.h', { h: 2 }),
	{
		...media('addMedia.link', { link: 'https://example.com/probe.mp4' }, { type: 'video' }),
		unwritten: 'https://github.com/shbernal/pptx-ts/issues/203',
	},
	media('addMedia.loop', { loop: true }),
	media('addMedia.loopCount', { loopCount: 3 }),
	media('addMedia.objectLock.noAdjustHandles', { objectLock: { noAdjustHandles: true } }),
	media('addMedia.objectLock.noChangeArrowheads', { objectLock: { noChangeArrowheads: true } }),
	media('addMedia.objectLock.noChangeAspect:false', { objectLock: { noChangeAspect: false } }, {}, [
		'addMedia.objectLock.noChangeAspect',
	]),
	media('addMedia.objectLock.noChangeShapeType', { objectLock: { noChangeShapeType: true } }),
	media('addMedia.objectLock.noCrop', { objectLock: { noCrop: true } }),
	{
		...media('addMedia.objectLock.noDrilldown', { objectLock: { noDrilldown: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'the writer warns that noDrilldown is not valid on a:picLocks and drops it: it is a graphic-frame/group lock, which a media picture does not have',
		},
	},
	media('addMedia.objectLock.noEditPoints', { objectLock: { noEditPoints: true } }),
	media('addMedia.objectLock.noGrp', { objectLock: { noGrp: true } }),
	media('addMedia.objectLock.noMove', { objectLock: { noMove: true } }),
	media('addMedia.objectLock.noResize', { objectLock: { noResize: true } }),
	media('addMedia.objectLock.noRot', { objectLock: { noRot: true } }),
	media('addMedia.objectLock.noSelect', { objectLock: { noSelect: true } }),
	{
		...media('addMedia.objectLock.noTextEdit', { objectLock: { noTextEdit: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'the writer warns that noTextEdit is not valid on a:picLocks and drops it: it is a shape lock (a:spLocks), and a media picture has no text body',
		},
	},
	{
		...media('addMedia.objectLock.noUngrp', { objectLock: { noUngrp: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'the writer drops noUngrp on a:picLocks: it is a group lock (a:grpSpLocks only), which a media picture does not have',
		},
	},
	media('addMedia.objectName', { objectName: 'Renamed probe' }),
	media('addMedia.type:video', { type: 'video' }, {}, ['addMedia.type']),
	media('addMedia.w', { w: 2 }),
	media('addMedia.x', { x: 2 }),
	media('addMedia.y', { y: 2 }),
	{
		...comment('addComment', {}, ['addComment.author', 'addComment.text']),
		// The comment itself survives since #221; it comes back renumbered (idx 1 -> 2).
		issue: 'https://github.com/shbernal/pptx-ts/issues/244',
	},
	comment('addComment.initials', { initials: 'XQ' }),
	comment('addComment.x', { x: 3 }),
	comment('addComment.y', { y: 4 }),
	comment('addComment.date', { date: '2026-01-02T03:04:05.000Z' }),
	animation('addAnimation.preset', { preset: 'fadeIn', objectName: 'probe' }, undefined, [
		'addAnimation.preset',
		'addAnimation.objectName',
	]),
	animation('addAnimation.preset:spin', { preset: 'spin' }, { preset: 'fadeIn', objectName: 'probe' }),
	animation(
		'addAnimation.trigger:afterPrevious',
		{ trigger: 'afterPrevious' },
		{ preset: 'fadeIn', objectName: 'probe' }
	),
	animation('addAnimation.durationMs', { durationMs: 1700 }, { preset: 'fadeIn', objectName: 'probe' }),
	chart('addChart.altText', { altText: 'A probe description' }),
	chart('addChart.chartColors', { chartColors: ['9B1B30', '250F6B'] }),
	chart('addChart.chartColorsOpacity', { chartColorsOpacity: 40 }),
	chart('addChart.dataBorder.color', { dataBorder: { color: '9B1B30' } }),
	chart('addChart.dataBorder.width', { dataBorder: { width: 3 } }),
	chart('addChart.dataBorder.dashType', { dataBorder: { dashType: 'lgDashDot' } }),
	chart('addChart.dataBorder.cap', { dataBorder: { cap: 'round' } }),
	chart('addChart.dataBorder.type:dash', { dataBorder: { type: 'dash' } }, ['addChart.dataBorder.type']),
	chartOn(
		'addChart.dataBorder.transparency',
		'bar',
		{ dataBorder: { color: '9B1B30' } },
		{ dataBorder: { color: '9B1B30', transparency: 40 } }
	),
	chartOn('addChart.displayBlanksAs:zero', 'line', {}, { displayBlanksAs: 'zero' }, ['addChart.displayBlanksAs']),
	chartWith('addChart.invertedColors', 'bar', NEGATIVE, {}, { invertedColors: ['9B1B30', '250F6B', '0B6E4F'] }),
	chart('addChart.lang', { lang: 'fr-FR' }),
	chart('addChart.layout', { layout: { x: 0.1, y: 0.15, w: 0.7, h: 0.6 } }, [
		'addChart.layout.x',
		'addChart.layout.y',
		'addChart.layout.w',
		'addChart.layout.h',
	]),
	chart('addChart.shadow.type:outer', { shadow: { type: 'outer' } }, ['addChart.shadow.type']),
	chartOn('addChart.shadow.angle', 'bar', { shadow: { type: 'outer' } }, { shadow: { type: 'outer', angle: 120 } }),
	chartOn('addChart.shadow.blur', 'bar', { shadow: { type: 'outer' } }, { shadow: { type: 'outer', blur: 6 } }),
	chartOn(
		'addChart.shadow.color',
		'bar',
		{ shadow: { type: 'outer' } },
		{ shadow: { type: 'outer', color: '9B1B30' } }
	),
	chartOn('addChart.shadow.offset', 'bar', { shadow: { type: 'outer' } }, { shadow: { type: 'outer', offset: 5 } }),
	chartOn(
		'addChart.shadow.rotateWithShape',
		'bar',
		{ shadow: { type: 'outer' } },
		{ shadow: { type: 'outer', rotateWithShape: false } }
	),
	chartOn(
		'addChart.shadow.transparency',
		'bar',
		{ shadow: { type: 'outer' } },
		{ shadow: { type: 'outer', transparency: 40 } }
	),
	chartWith('addChart.showBubbleSize', 'bubble', BUBBLES, {}, { showBubbleSize: true }),
	chartOn('addChart.showLabel', 'pie', {}, { showLabel: true }),
	chartOn('addChart.showLeaderLines', 'pie', { showValue: true }, { showLeaderLines: true }),
	chartOn('addChart.leaderLineColor', 'pie', { showValue: true, showLeaderLines: true }, { leaderLineColor: '9B1B30' }),
	chartOn('addChart.leaderLineSize', 'pie', { showValue: true, showLeaderLines: true }, { leaderLineSize: 2.5 }),
	chart('addChart.showLegend', { showLegend: true }),
	chartOn('addChart.showPercent', 'pie', {}, { showPercent: true }),
	chart('addChart.showSerName', { showSerName: true }),
	chartOn('addChart.bar3DShape', 'bar3D', {}, { bar3DShape: 'cylinder' }),
	chart('addChart.barDir', { barDir: 'bar' }),
	chartOn('addChart.barGapDepthPct', 'bar3D', {}, { barGapDepthPct: 250 }),
	chart('addChart.barGapWidthPct', { barGapWidthPct: 60 }),
	chart('addChart.barGrouping', { barGrouping: 'stacked' }),
	chart('addChart.barOverlapPct', { barOverlapPct: 40 }),
	chartOn('addChart.barSeriesLine', 'bar', { barGrouping: 'stacked' }, { barSeriesLine: true }),
	chartOn(
		'addChart.barSeriesLine.cap',
		'bar',
		{ barGrouping: 'stacked', barSeriesLine: { color: '777777' } },
		{ barSeriesLine: { color: '777777', cap: 'round' } }
	),
	chartOn(
		'addChart.barSeriesLine.color',
		'bar',
		{ barGrouping: 'stacked', barSeriesLine: { size: 1 } },
		{ barSeriesLine: { size: 1, color: '9B1B30' } }
	),
	chartOn(
		'addChart.barSeriesLine.dashType',
		'bar',
		{ barGrouping: 'stacked', barSeriesLine: { color: '777777' } },
		{ barSeriesLine: { color: '777777', dashType: 'lgDashDot' } }
	),
	chartOn(
		'addChart.barSeriesLine.size',
		'bar',
		{ barGrouping: 'stacked', barSeriesLine: { color: '777777' } },
		{ barSeriesLine: { color: '777777', size: 2.5 } }
	),
	chartOn(
		'addChart.barSeriesLine.style',
		'bar',
		{ barGrouping: 'stacked', barSeriesLine: { color: '777777' } },
		{ barSeriesLine: { color: '777777', style: 'dash' } }
	),
	chartOn(
		'addChart.barSeriesLine.transparency',
		'bar',
		{ barGrouping: 'stacked', barSeriesLine: { color: '777777' } },
		{ barSeriesLine: { color: '777777', transparency: 40 } }
	),
	chartOn(
		'addChart.barSeriesLine.type',
		'bar',
		{ barGrouping: 'stacked', barSeriesLine: { color: '777777' } },
		{ barSeriesLine: { color: '777777', type: 'dash' } }
	),
	chartOn(
		'addChart.barSeriesLine.width',
		'bar',
		{ barGrouping: 'stacked', barSeriesLine: { color: '777777' } },
		{ barSeriesLine: { color: '777777', width: 3 } }
	),
	chartOn(
		'addChart.catGridLine.cap',
		'bar',
		{ catGridLine: { color: '777777' } },
		{ catGridLine: { color: '777777', cap: 'round' } }
	),
	chartOn(
		'addChart.catGridLine.color',
		'bar',
		{ catGridLine: { width: 2 } },
		{ catGridLine: { width: 2, color: '9B1B30' } }
	),
	chartOn(
		'addChart.catGridLine.dashType',
		'bar',
		{ catGridLine: { color: '777777' } },
		{ catGridLine: { color: '777777', dashType: 'lgDashDot' } }
	),
	chartOn(
		'addChart.catGridLine.size',
		'bar',
		{ catGridLine: { color: '777777' } },
		{ catGridLine: { color: '777777', size: 2.5 } }
	),
	chartOn(
		'addChart.catGridLine.style',
		'bar',
		{ catGridLine: { color: '777777' } },
		{ catGridLine: { color: '777777', style: 'dash' } }
	),
	chartOn(
		'addChart.catGridLine.transparency',
		'bar',
		{ catGridLine: { color: '777777' } },
		{ catGridLine: { color: '777777', transparency: 40 } }
	),
	chartOn(
		'addChart.catGridLine.type',
		'bar',
		{ catGridLine: { color: '777777' } },
		{ catGridLine: { color: '777777', type: 'dash' } }
	),
	chartOn(
		'addChart.catGridLine.width',
		'bar',
		{ catGridLine: { color: '777777' } },
		{ catGridLine: { color: '777777', width: 3 } }
	),
	chartOn(
		'addChart.valGridLine.cap',
		'bar',
		{ valGridLine: { color: '777777' } },
		{ valGridLine: { color: '777777', cap: 'round' } }
	),
	chartOn(
		'addChart.valGridLine.color',
		'bar',
		{ valGridLine: { width: 2 } },
		{ valGridLine: { width: 2, color: '9B1B30' } }
	),
	chartOn(
		'addChart.valGridLine.dashType',
		'bar',
		{ valGridLine: { color: '777777' } },
		{ valGridLine: { color: '777777', dashType: 'lgDashDot' } }
	),
	chartOn(
		'addChart.valGridLine.size',
		'bar',
		{ valGridLine: { color: '777777' } },
		{ valGridLine: { color: '777777', size: 2.5 } }
	),
	chartOn(
		'addChart.valGridLine.style',
		'bar',
		{ valGridLine: { color: '777777' } },
		{ valGridLine: { color: '777777', style: 'dash' } }
	),
	chartOn(
		'addChart.valGridLine.transparency',
		'bar',
		{ valGridLine: { color: '777777' } },
		{ valGridLine: { color: '777777', transparency: 40 } }
	),
	chartOn(
		'addChart.valGridLine.type',
		'bar',
		{ valGridLine: { color: '777777' } },
		{ valGridLine: { color: '777777', type: 'dash' } }
	),
	chartOn(
		'addChart.valGridLine.width',
		'bar',
		{ valGridLine: { color: '777777' } },
		{ valGridLine: { color: '777777', width: 3 } }
	),
	chartOn(
		'addChart.serGridLine.cap',
		'bar3D',
		{ barGrouping: 'standard', serGridLine: { color: '777777' } },
		{ serGridLine: { color: '777777', cap: 'round' } }
	),
	chartOn(
		'addChart.serGridLine.color',
		'bar3D',
		{ barGrouping: 'standard', serGridLine: { width: 2 } },
		{ serGridLine: { width: 2, color: '9B1B30' } }
	),
	chartOn(
		'addChart.serGridLine.dashType',
		'bar3D',
		{ barGrouping: 'standard', serGridLine: { color: '777777' } },
		{ serGridLine: { color: '777777', dashType: 'lgDashDot' } }
	),
	chartOn(
		'addChart.serGridLine.size',
		'bar3D',
		{ barGrouping: 'standard', serGridLine: { color: '777777' } },
		{ serGridLine: { color: '777777', size: 2.5 } }
	),
	chartOn(
		'addChart.serGridLine.style',
		'bar3D',
		{ barGrouping: 'standard', serGridLine: { color: '777777' } },
		{ serGridLine: { color: '777777', style: 'dash' } }
	),
	chartOn(
		'addChart.serGridLine.transparency',
		'bar3D',
		{ barGrouping: 'standard', serGridLine: { color: '777777' } },
		{ serGridLine: { color: '777777', transparency: 40 } }
	),
	chartOn(
		'addChart.serGridLine.type',
		'bar3D',
		{ barGrouping: 'standard', serGridLine: { color: '777777' } },
		{ serGridLine: { color: '777777', type: 'dash' } }
	),
	chartOn(
		'addChart.serGridLine.width',
		'bar3D',
		{ barGrouping: 'standard', serGridLine: { color: '777777' } },
		{ serGridLine: { color: '777777', width: 3 } }
	),
	chartOn(
		'addChart.catAxisLine.cap',
		'bar',
		{ catAxisLine: { color: '777777' } },
		{ catAxisLine: { color: '777777', cap: 'round' } }
	),
	chartOn(
		'addChart.catAxisLine.color',
		'bar',
		{ catAxisLine: { width: 2 } },
		{ catAxisLine: { width: 2, color: '9B1B30' } }
	),
	chartOn(
		'addChart.catAxisLine.dashType',
		'bar',
		{ catAxisLine: { color: '777777' } },
		{ catAxisLine: { color: '777777', dashType: 'lgDashDot' } }
	),
	chartOn(
		'addChart.catAxisLine.transparency',
		'bar',
		{ catAxisLine: { color: '777777' } },
		{ catAxisLine: { color: '777777', transparency: 40 } }
	),
	chartOn(
		'addChart.catAxisLine.type',
		'bar',
		{ catAxisLine: { color: '777777' } },
		{ catAxisLine: { color: '777777', type: 'dash' } }
	),
	chartOn(
		'addChart.catAxisLine.width',
		'bar',
		{ catAxisLine: { color: '777777' } },
		{ catAxisLine: { color: '777777', width: 3 } }
	),
	chartOn(
		'addChart.valAxisLine.cap',
		'bar',
		{ valAxisLine: { color: '777777' } },
		{ valAxisLine: { color: '777777', cap: 'round' } }
	),
	chartOn(
		'addChart.valAxisLine.color',
		'bar',
		{ valAxisLine: { width: 2 } },
		{ valAxisLine: { width: 2, color: '9B1B30' } }
	),
	chartOn(
		'addChart.valAxisLine.dashType',
		'bar',
		{ valAxisLine: { color: '777777' } },
		{ valAxisLine: { color: '777777', dashType: 'lgDashDot' } }
	),
	chartOn(
		'addChart.valAxisLine.transparency',
		'bar',
		{ valAxisLine: { color: '777777' } },
		{ valAxisLine: { color: '777777', transparency: 40 } }
	),
	chartOn(
		'addChart.valAxisLine.type',
		'bar',
		{ valAxisLine: { color: '777777' } },
		{ valAxisLine: { color: '777777', type: 'dash' } }
	),
	chartOn(
		'addChart.valAxisLine.width',
		'bar',
		{ valAxisLine: { color: '777777' } },
		{ valAxisLine: { color: '777777', width: 3 } }
	),
	chartOn(
		'addChart.serAxisLine.cap',
		'bar3D',
		{ barGrouping: 'standard', serAxisLine: { color: '777777' } },
		{ serAxisLine: { color: '777777', cap: 'round' } }
	),
	chartOn(
		'addChart.serAxisLine.color',
		'bar3D',
		{ barGrouping: 'standard', serAxisLine: { width: 2 } },
		{ serAxisLine: { width: 2, color: '9B1B30' } }
	),
	chartOn(
		'addChart.serAxisLine.dashType',
		'bar3D',
		{ barGrouping: 'standard', serAxisLine: { color: '777777' } },
		{ serAxisLine: { color: '777777', dashType: 'lgDashDot' } }
	),
	chartOn(
		'addChart.serAxisLine.transparency',
		'bar3D',
		{ barGrouping: 'standard', serAxisLine: { color: '777777' } },
		{ serAxisLine: { color: '777777', transparency: 40 } }
	),
	chartOn(
		'addChart.serAxisLine.type',
		'bar3D',
		{ barGrouping: 'standard', serAxisLine: { color: '777777' } },
		{ serAxisLine: { color: '777777', type: 'dash' } }
	),
	chartOn(
		'addChart.serAxisLine.width',
		'bar3D',
		{ barGrouping: 'standard', serAxisLine: { color: '777777' } },
		{ serAxisLine: { color: '777777', width: 3 } }
	),
	chart('addChart.catAxisLineColor', { catAxisLineColor: '9B1B30' }),
	chart('addChart.catAxisLineShow', { catAxisLineShow: false }),
	chart('addChart.catAxisLineSize', { catAxisLineSize: 3 }),
	chart('addChart.catAxisLineStyle', { catAxisLineStyle: 'dot' }),
	chart('addChart.valAxisLineColor', { valAxisLineColor: '9B1B30' }),
	chart('addChart.valAxisLineShow', { valAxisLineShow: false }),
	chart('addChart.valAxisLineSize', { valAxisLineSize: 3 }),
	chart('addChart.valAxisLineStyle', { valAxisLineStyle: 'dot' }),
	chartOn('addChart.serAxisLineColor', 'bar3D', { barGrouping: 'standard' }, { serAxisLineColor: '9B1B30' }),
	chartOn('addChart.serAxisLineShow', 'bar3D', { barGrouping: 'standard' }, { serAxisLineShow: false }),
	chartOn('addChart.serAxisLineSize', 'bar3D', { barGrouping: 'standard' }, { serAxisLineSize: 3 }),
	chartOn('addChart.serAxisLineStyle', 'bar3D', { barGrouping: 'standard' }, { serAxisLineStyle: 'dot' }),
	chartOn('addChart.showTitle', 'bar', { title: 'Probe title' }, { showTitle: true }),
	chartOn('addChart.title', 'bar', { showTitle: true }, { title: 'Probe title' }),
	chart('addChart.showValue', { showValue: true }),
	chartOn('addChart.v3DPerspective', 'bar3D', {}, { v3DPerspective: 60 }),
	chartOn('addChart.v3DRAngAx', 'bar3D', {}, { v3DRAngAx: true }),
	chartOn('addChart.v3DRotX', 'bar3D', {}, { v3DRotX: 45 }),
	chartOn('addChart.v3DRotY', 'bar3D', {}, { v3DRotY: 60 }),
	chart('addChart.chartArea.fill.color', { chartArea: { fill: { color: 'F0C808' } } }, [
		'addChart.chartArea.fill',
		'addChart.chartArea.fill.color',
	]),
	chartOn(
		'addChart.chartArea.fill.transparency',
		'bar',
		{ chartArea: { fill: { color: 'F0C808' } } },
		{ chartArea: { fill: { color: 'F0C808', transparency: 40 } } }
	),
	chart('addChart.chartArea.fill.type:inherit', { chartArea: { fill: { type: 'inherit' } } }, [
		'addChart.chartArea.fill.type',
	]),
	chart(
		'addChart.chartArea.fill.gradient',
		{
			chartArea: {
				fill: {
					type: 'gradient',
					gradient: {
						kind: 'linear',
						angle: 45,
						stops: [
							{ position: 0, color: '9B1B30' },
							{ position: 100, color: '250F6B', transparency: 30 },
						],
					},
				},
			},
		},
		[
			'addChart.chartArea.fill.gradient.kind',
			'addChart.chartArea.fill.gradient.angle',
			'addChart.chartArea.fill.gradient.stops.color',
			'addChart.chartArea.fill.gradient.stops.position',
			'addChart.chartArea.fill.gradient.stops.transparency',
		]
	),
	chartOn(
		'addChart.chartArea.fill.gradient.rotateWithShape',
		'bar',
		{
			chartArea: {
				fill: {
					type: 'gradient',
					gradient: {
						kind: 'linear',
						angle: 45,
						stops: [
							{ position: 0, color: '9B1B30' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			},
		},
		{
			chartArea: {
				fill: {
					type: 'gradient',
					gradient: {
						kind: 'linear',
						angle: 45,
						stops: [
							{ position: 0, color: '9B1B30' },
							{ position: 100, color: '250F6B' },
						],
						rotateWithShape: false,
					},
				},
			},
		}
	),
	chartOn(
		'addChart.chartArea.fill.gradient.scaled',
		'bar',
		{
			chartArea: {
				fill: {
					type: 'gradient',
					gradient: {
						kind: 'linear',
						angle: 45,
						stops: [
							{ position: 0, color: '9B1B30' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			},
		},
		{
			chartArea: {
				fill: {
					type: 'gradient',
					gradient: {
						kind: 'linear',
						angle: 45,
						stops: [
							{ position: 0, color: '9B1B30' },
							{ position: 100, color: '250F6B' },
						],
						scaled: true,
					},
				},
			},
		}
	),
	chart(
		'addChart.chartArea.fill.pattern',
		{ chartArea: { fill: { type: 'pattern', pattern: { preset: 'dkDnDiag', fgColor: '9B1B30', bgColor: 'F0C808' } } } },
		[
			'addChart.chartArea.fill.pattern.preset',
			'addChart.chartArea.fill.pattern.fgColor',
			'addChart.chartArea.fill.pattern.bgColor',
		]
	),
	chart('addChart.chartArea.border.color', { chartArea: { border: { color: '9B1B30' } } }),
	chart('addChart.chartArea.border.width', { chartArea: { border: { width: 3 } } }),
	chart('addChart.chartArea.border.dashType', { chartArea: { border: { dashType: 'lgDashDot' } } }),
	chart('addChart.chartArea.border.cap', { chartArea: { border: { cap: 'round' } } }),
	chart('addChart.chartArea.border.type:dash', { chartArea: { border: { type: 'dash' } } }, [
		'addChart.chartArea.border.type',
	]),
	chartOn(
		'addChart.chartArea.border.transparency',
		'bar',
		{ chartArea: { border: { color: '9B1B30' } } },
		{ chartArea: { border: { color: '9B1B30', transparency: 40 } } }
	),
	chart('addChart.plotArea.fill.color', { plotArea: { fill: { color: 'F0C808' } } }, [
		'addChart.plotArea.fill',
		'addChart.plotArea.fill.color',
	]),
	chartOn(
		'addChart.plotArea.fill.transparency',
		'bar',
		{ plotArea: { fill: { color: 'F0C808' } } },
		{ plotArea: { fill: { color: 'F0C808', transparency: 40 } } }
	),
	chart('addChart.plotArea.fill.type:inherit', { plotArea: { fill: { type: 'inherit' } } }, [
		'addChart.plotArea.fill.type',
	]),
	chart(
		'addChart.plotArea.fill.gradient',
		{
			plotArea: {
				fill: {
					type: 'gradient',
					gradient: {
						kind: 'linear',
						angle: 45,
						stops: [
							{ position: 0, color: '9B1B30' },
							{ position: 100, color: '250F6B', transparency: 30 },
						],
					},
				},
			},
		},
		[
			'addChart.plotArea.fill.gradient.kind',
			'addChart.plotArea.fill.gradient.angle',
			'addChart.plotArea.fill.gradient.stops.color',
			'addChart.plotArea.fill.gradient.stops.position',
			'addChart.plotArea.fill.gradient.stops.transparency',
		]
	),
	chartOn(
		'addChart.plotArea.fill.gradient.rotateWithShape',
		'bar',
		{
			plotArea: {
				fill: {
					type: 'gradient',
					gradient: {
						kind: 'linear',
						angle: 45,
						stops: [
							{ position: 0, color: '9B1B30' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			},
		},
		{
			plotArea: {
				fill: {
					type: 'gradient',
					gradient: {
						kind: 'linear',
						angle: 45,
						stops: [
							{ position: 0, color: '9B1B30' },
							{ position: 100, color: '250F6B' },
						],
						rotateWithShape: false,
					},
				},
			},
		}
	),
	chartOn(
		'addChart.plotArea.fill.gradient.scaled',
		'bar',
		{
			plotArea: {
				fill: {
					type: 'gradient',
					gradient: {
						kind: 'linear',
						angle: 45,
						stops: [
							{ position: 0, color: '9B1B30' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			},
		},
		{
			plotArea: {
				fill: {
					type: 'gradient',
					gradient: {
						kind: 'linear',
						angle: 45,
						stops: [
							{ position: 0, color: '9B1B30' },
							{ position: 100, color: '250F6B' },
						],
						scaled: true,
					},
				},
			},
		}
	),
	chart(
		'addChart.plotArea.fill.pattern',
		{ plotArea: { fill: { type: 'pattern', pattern: { preset: 'dkDnDiag', fgColor: '9B1B30', bgColor: 'F0C808' } } } },
		[
			'addChart.plotArea.fill.pattern.preset',
			'addChart.plotArea.fill.pattern.fgColor',
			'addChart.plotArea.fill.pattern.bgColor',
		]
	),
	chart('addChart.plotArea.border.color', { plotArea: { border: { color: '9B1B30' } } }),
	chart('addChart.plotArea.border.width', { plotArea: { border: { width: 3 } } }),
	chart('addChart.plotArea.border.dashType', { plotArea: { border: { dashType: 'lgDashDot' } } }),
	chart('addChart.plotArea.border.cap', { plotArea: { border: { cap: 'round' } } }),
	chart('addChart.plotArea.border.type:dash', { plotArea: { border: { type: 'dash' } } }, [
		'addChart.plotArea.border.type',
	]),
	chartOn(
		'addChart.plotArea.border.transparency',
		'bar',
		{ plotArea: { border: { color: '9B1B30' } } },
		{ plotArea: { border: { color: '9B1B30', transparency: 40 } } }
	),
	chartOn(
		'addChart.chartArea.roundedCorners',
		'bar',
		{ chartArea: { fill: { color: 'F0C808' } } },
		{ chartArea: { fill: { color: 'F0C808' }, roundedCorners: false } }
	),
	chartOn(
		'addChart.seriesOptions.color',
		'bar',
		{ showValue: true },
		{ showValue: true, seriesOptions: [{ color: '9B1B30' }] }
	),
	chartOn(
		'addChart.seriesOptions.dataLabelColor',
		'bar',
		{ showValue: true },
		{ showValue: true, seriesOptions: [{ dataLabelColor: '9B1B30' }] }
	),
	chartOn(
		'addChart.seriesOptions.dataLabelFontBold',
		'bar',
		{ showValue: true },
		{ showValue: true, seriesOptions: [{ dataLabelFontBold: true }] }
	),
	chartOn(
		'addChart.seriesOptions.dataLabelFontFace',
		'bar',
		{ showValue: true },
		{ showValue: true, seriesOptions: [{ dataLabelFontFace: 'Georgia' }] }
	),
	chartOn(
		'addChart.seriesOptions.dataLabelFontItalic',
		'bar',
		{ showValue: true },
		{ showValue: true, seriesOptions: [{ dataLabelFontItalic: true }] }
	),
	chartOn(
		'addChart.seriesOptions.dataLabelFontSize',
		'bar',
		{ showValue: true },
		{ showValue: true, seriesOptions: [{ dataLabelFontSize: 18 }] }
	),
	chartOn(
		'addChart.seriesOptions.dataLabelFormatCode',
		'bar',
		{ showValue: true },
		{ showValue: true, seriesOptions: [{ dataLabelFormatCode: '0.00' }] }
	),
	chartOn('addChart.seriesOptions.lineSize', 'line', {}, { seriesOptions: [{ lineSize: 5 }] }),
	chartOn('addChart.dataLabelBkgrdColors', 'bar', { showValue: true }, { dataLabelBkgrdColors: true }),
	chartOn('addChart.dataLabelColor', 'bar', { showValue: true }, { dataLabelColor: '9B1B30' }),
	chartOn('addChart.dataLabelFontBold', 'bar', { showValue: true }, { dataLabelFontBold: true }),
	chartOn('addChart.dataLabelFontFace', 'bar', { showValue: true }, { dataLabelFontFace: 'Georgia' }),
	chartOn('addChart.dataLabelFontItalic', 'bar', { showValue: true }, { dataLabelFontItalic: true }),
	chartOn('addChart.dataLabelFontSize', 'bar', { showValue: true }, { dataLabelFontSize: 18 }),
	chartOn('addChart.dataLabelFormatCode', 'bar', { showValue: true }, { dataLabelFormatCode: '0.00' }),
	chartOn('addChart.dataLabelPosition:inEnd', 'bar', { showValue: true }, { dataLabelPosition: 'inEnd' }, [
		'addChart.dataLabelPosition',
	]),
	chartOn('addChart.dataLabelFormatScatter:XY', 'scatter', { showLabel: true }, { dataLabelFormatScatter: 'XY' }, [
		'addChart.dataLabelFormatScatter',
	]),
	chart('addChart.showDataTable', { showDataTable: true }),
	chartOn('addChart.dataTableFontSize', 'bar', { showDataTable: true }, { dataTableFontSize: 16 }),
	chartOn('addChart.dataTableFormatCode', 'bar', { showDataTable: true }, { dataTableFormatCode: '0.00' }),
	chartOn('addChart.showDataTableHorzBorder', 'bar', { showDataTable: true }, { showDataTableHorzBorder: true }),
	chartOn('addChart.showDataTableKeys', 'bar', { showDataTable: true }, { showDataTableKeys: true }),
	chartOn('addChart.showDataTableOutline', 'bar', { showDataTable: true }, { showDataTableOutline: true }),
	chartOn('addChart.showDataTableVertBorder', 'bar', { showDataTable: true }, { showDataTableVertBorder: true }),
	chartOn('addChart.legendColor', 'bar', { showLegend: true }, { legendColor: '9B1B30' }),
	chartOn('addChart.legendFontFace', 'bar', { showLegend: true }, { legendFontFace: 'Georgia' }),
	chartOn('addChart.legendFontSize', 'bar', { showLegend: true }, { legendFontSize: 18 }),
	chartOn('addChart.legendPos:t', 'bar', { showLegend: true }, { legendPos: 't' }, ['addChart.legendPos']),
	chartOn('addChart.legendLayout', 'bar', { showLegend: true }, { legendLayout: { x: 0.7, y: 0.3, w: 0.25, h: 0.4 } }, [
		'addChart.legendLayout.x',
		'addChart.legendLayout.y',
		'addChart.legendLayout.w',
		'addChart.legendLayout.h',
	]),
	chartOn('addChart.titleAlign', 'bar', { showTitle: true, title: 'Probe title' }, { titleAlign: 'left' }),
	chartOn('addChart.titleBold', 'bar', { showTitle: true, title: 'Probe title' }, { titleBold: true }),
	chartOn('addChart.titleColor', 'bar', { showTitle: true, title: 'Probe title' }, { titleColor: '9B1B30' }),
	chartOn('addChart.titleFontFace', 'bar', { showTitle: true, title: 'Probe title' }, { titleFontFace: 'Georgia' }),
	chartOn('addChart.titleFontSize', 'bar', { showTitle: true, title: 'Probe title' }, { titleFontSize: 28 }),
	chartOn('addChart.titleItalic', 'bar', { showTitle: true, title: 'Probe title' }, { titleItalic: true }),
	chartOn('addChart.titleRotate', 'bar', { showTitle: true, title: 'Probe title' }, { titleRotate: 20 }),
	chartOn('addChart.titleUnderline', 'bar', { showTitle: true, title: 'Probe title' }, { titleUnderline: true }),
	chartOn('addChart.titlePos', 'bar', { showTitle: true, title: 'Probe title' }, { titlePos: { x: 0.5, y: 0.3 } }, [
		'addChart.titlePos.x',
		'addChart.titlePos.y',
	]),
	chartOn('addChart.lineCap', 'line', {}, { lineCap: 'round' }),
	chartOn('addChart.lineDash:dash', 'line', {}, { lineDash: 'dash' }, ['addChart.lineDash']),
	chartOn('addChart.lineDashValues', 'line', {}, { lineDashValues: ['lgDash', 'sysDot'] }),
	chartOn('addChart.lineDataSymbol:diamond', 'line', {}, { lineDataSymbol: 'diamond' }, ['addChart.lineDataSymbol']),
	chartOn('addChart.lineDataSymbolLineColor', 'line', {}, { lineDataSymbolLineColor: '9B1B30' }),
	chartOn('addChart.lineDataSymbolLineSize', 'line', {}, { lineDataSymbolLineSize: 3 }),
	chartOn('addChart.lineDataSymbolSize', 'line', {}, { lineDataSymbolSize: 12 }),
	chartOn('addChart.lineSize', 'line', {}, { lineSize: 5 }),
	chartOn('addChart.lineSmooth', 'line', {}, { lineSmooth: true }),
	chartOn('addChart.firstSliceAng', 'pie', {}, { firstSliceAng: 90 }),
	chartOn('addChart.dataNoEffects', 'pie', { shadow: { type: 'outer' } }, { dataNoEffects: true }),
	chartOn('addChart.holeSize', 'doughnut', {}, { holeSize: 75 }),
	chartOn('addChart.radarStyle:filled', 'radar', {}, { radarStyle: 'filled' }, ['addChart.radarStyle']),
	chart('addChart.metadata', { metadata: { sourceId: 'probe' } }),
	chart('addChart.objectName', { objectName: 'renamed' }),
	chart('addChart.catAxisHidden', { catAxisHidden: true }),
	chart('addChart.catAxisCrossesAt', { catAxisCrossesAt: 2 }),
	chart('addChart.catAxisLabelColor', { catAxisLabelColor: '9B1B30' }),
	chart('addChart.catAxisLabelFontBold', { catAxisLabelFontBold: true }),
	chart('addChart.catAxisLabelFontFace', { catAxisLabelFontFace: 'Georgia' }),
	chart('addChart.catAxisLabelFontItalic', { catAxisLabelFontItalic: true }),
	chart('addChart.catAxisLabelFontSize', { catAxisLabelFontSize: 18 }),
	chart('addChart.catAxisLabelFrequency', { catAxisLabelFrequency: 2 }),
	chart('addChart.catAxisLabelRotate', { catAxisLabelRotate: 45 }),
	chart('addChart.catAxisLabelPos:high', { catAxisLabelPos: 'high' }, ['addChart.catAxisLabelPos']),
	chart('addChart.catAxisMajorTickMark:cross', { catAxisMajorTickMark: 'cross' }, ['addChart.catAxisMajorTickMark']),
	chart('addChart.catAxisMinorTickMark:in', { catAxisMinorTickMark: 'in' }, ['addChart.catAxisMinorTickMark']),
	chart('addChart.catAxisOrientation:maxMin', { catAxisOrientation: 'maxMin' }, ['addChart.catAxisOrientation']),
	chartOn('addChart.catAxisLabelFormatCode', 'scatter', {}, { catAxisLabelFormatCode: '0.00' }),
	chartOn('addChart.catAxisMajorUnit', 'scatter', {}, { catAxisMajorUnit: 2 }),
	chartOn('addChart.catAxisMinorUnit', 'scatter', {}, { catAxisMinorUnit: 0.5 }),
	chartOn('addChart.catAxisMaxVal', 'scatter', {}, { catAxisMaxVal: 10 }),
	chartOn('addChart.catAxisMinVal', 'scatter', {}, { catAxisMinVal: 1 }),
	chartWith('addChart.catLabelFormatCode', 'bar', DATED, {}, { catLabelFormatCode: 'yyyy-mm-dd' }),
	chartWith(
		'addChart.catAxisBaseTimeUnit:months',
		'bar',
		DATED,
		{ catLabelFormatCode: 'yyyy-mm-dd' },
		{ catAxisBaseTimeUnit: 'months' },
		['addChart.catAxisBaseTimeUnit']
	),
	chartWith(
		'addChart.catAxisMajorTimeUnit:months',
		'bar',
		DATED,
		{ catLabelFormatCode: 'yyyy-mm-dd' },
		{ catAxisMajorTimeUnit: 'months' },
		['addChart.catAxisMajorTimeUnit']
	),
	chartWith(
		'addChart.catAxisMinorTimeUnit:days',
		'bar',
		DATED,
		{ catLabelFormatCode: 'yyyy-mm-dd' },
		{ catAxisMinorTimeUnit: 'days' },
		['addChart.catAxisMinorTimeUnit']
	),
	chartWith('addChart.catAxisMultiLevelLabels', 'bar', LEVELS, {}, { catAxisMultiLevelLabels: true }),
	chartOn('addChart.showCatAxisTitle', 'bar', { catAxisTitle: 'Category' }, { showCatAxisTitle: true }),
	chartOn('addChart.catAxisTitle', 'bar', { showCatAxisTitle: true }, { catAxisTitle: 'Category' }),
	chartOn(
		'addChart.catAxisTitleColor',
		'bar',
		{ showCatAxisTitle: true, catAxisTitle: 'Category' },
		{ catAxisTitleColor: '9B1B30' }
	),
	chartOn(
		'addChart.catAxisTitleFontFace',
		'bar',
		{ showCatAxisTitle: true, catAxisTitle: 'Category' },
		{ catAxisTitleFontFace: 'Georgia' }
	),
	chartOn(
		'addChart.catAxisTitleFontSize',
		'bar',
		{ showCatAxisTitle: true, catAxisTitle: 'Category' },
		{ catAxisTitleFontSize: 20 }
	),
	chartOn(
		'addChart.catAxisTitleRotate',
		'bar',
		{ showCatAxisTitle: true, catAxisTitle: 'Category' },
		{ catAxisTitleRotate: 45 }
	),
	chart('addChart.valAxisHidden', { valAxisHidden: true }),
	chart('addChart.valAxisCrossesAt', { valAxisCrossesAt: 2 }),
	chartOn(
		'addChart.valAxisDisplayUnitLabel',
		'bar',
		{ valAxisDisplayUnit: 'thousands' },
		{ valAxisDisplayUnitLabel: true }
	),
	chart('addChart.valAxisLabelColor', { valAxisLabelColor: '9B1B30' }),
	chart('addChart.valAxisLabelFontBold', { valAxisLabelFontBold: true }),
	chart('addChart.valAxisLabelFontFace', { valAxisLabelFontFace: 'Georgia' }),
	chart('addChart.valAxisLabelFontItalic', { valAxisLabelFontItalic: true }),
	chart('addChart.valAxisLabelFontSize', { valAxisLabelFontSize: 18 }),
	chart('addChart.valAxisLabelFormatCode', { valAxisLabelFormatCode: '0.00' }),
	chart('addChart.valAxisLabelRotate', { valAxisLabelRotate: 45 }),
	chart('addChart.valAxisLogScaleBase', { valAxisLogScaleBase: 10 }),
	chart('addChart.valAxisMajorUnit', { valAxisMajorUnit: 2 }),
	chart('addChart.valAxisMaxVal', { valAxisMaxVal: 12 }),
	chart('addChart.valAxisMinorUnit', { valAxisMinorUnit: 0.5 }),
	chart('addChart.valAxisMinVal', { valAxisMinVal: 1 }),
	chart('addChart.valLabelFormatCode', { valLabelFormatCode: '0.0%' }),
	chart('addChart.valAxisCrossBetween:midCat', { valAxisCrossBetween: 'midCat' }, ['addChart.valAxisCrossBetween']),
	chart('addChart.valAxisDisplayUnit:thousands', { valAxisDisplayUnit: 'thousands' }, ['addChart.valAxisDisplayUnit']),
	chart('addChart.valAxisLabelPos:high', { valAxisLabelPos: 'high' }, ['addChart.valAxisLabelPos']),
	chart('addChart.valAxisMajorTickMark:cross', { valAxisMajorTickMark: 'cross' }, ['addChart.valAxisMajorTickMark']),
	chart('addChart.valAxisMinorTickMark:in', { valAxisMinorTickMark: 'in' }, ['addChart.valAxisMinorTickMark']),
	chart('addChart.valAxisOrientation:maxMin', { valAxisOrientation: 'maxMin' }, ['addChart.valAxisOrientation']),
	chartOn('addChart.showValAxisTitle', 'bar', { valAxisTitle: 'Value' }, { showValAxisTitle: true }),
	chartOn('addChart.valAxisTitle', 'bar', { showValAxisTitle: true }, { valAxisTitle: 'Value' }),
	chartOn(
		'addChart.valAxisTitleColor',
		'bar',
		{ showValAxisTitle: true, valAxisTitle: 'Value' },
		{ valAxisTitleColor: '9B1B30' }
	),
	chartOn(
		'addChart.valAxisTitleFontFace',
		'bar',
		{ showValAxisTitle: true, valAxisTitle: 'Value' },
		{ valAxisTitleFontFace: 'Georgia' }
	),
	chartOn(
		'addChart.valAxisTitleFontSize',
		'bar',
		{ showValAxisTitle: true, valAxisTitle: 'Value' },
		{ valAxisTitleFontSize: 20 }
	),
	chartOn(
		'addChart.valAxisTitleRotate',
		'bar',
		{ showValAxisTitle: true, valAxisTitle: 'Value' },
		{ valAxisTitleRotate: 45 }
	),
	chartOn('addChart.serAxisHidden', 'bar3D', {}, { serAxisHidden: true }),
	chartOn('addChart.serAxisLabelColor', 'bar3D', {}, { serAxisLabelColor: '9B1B30' }),
	chartOn('addChart.serAxisLabelFontBold', 'bar3D', {}, { serAxisLabelFontBold: true }),
	chartOn('addChart.serAxisLabelFontFace', 'bar3D', {}, { serAxisLabelFontFace: 'Georgia' }),
	chartOn('addChart.serAxisLabelFontItalic', 'bar3D', {}, { serAxisLabelFontItalic: true }),
	chartOn('addChart.serAxisLabelFontSize', 'bar3D', {}, { serAxisLabelFontSize: 18 }),
	chartOn('addChart.serAxisLabelFrequency', 'bar3D', {}, { serAxisLabelFrequency: 2 }),
	chartOn('addChart.serLabelFormatCode', 'bar3D', {}, { serLabelFormatCode: '0.00' }),
	chartOn('addChart.serAxisLabelPos:high', 'bar3D', {}, { serAxisLabelPos: 'high' }, ['addChart.serAxisLabelPos']),
	chartOn('addChart.serAxisOrientation:maxMin', 'bar3D', {}, { serAxisOrientation: 'maxMin' }, [
		'addChart.serAxisOrientation',
	]),
	chartOn('addChart.showSerAxisTitle', 'bar3D', { serAxisTitle: 'Series' }, { showSerAxisTitle: true }),
	chartOn('addChart.serAxisTitle', 'bar3D', { showSerAxisTitle: true }, { serAxisTitle: 'Series' }),
	chartOn(
		'addChart.serAxisTitleColor',
		'bar3D',
		{ showSerAxisTitle: true, serAxisTitle: 'Series' },
		{ serAxisTitleColor: '9B1B30' }
	),
	chartOn(
		'addChart.serAxisTitleFontFace',
		'bar3D',
		{ showSerAxisTitle: true, serAxisTitle: 'Series' },
		{ serAxisTitleFontFace: 'Georgia' }
	),
	chartOn(
		'addChart.serAxisTitleFontSize',
		'bar3D',
		{ showSerAxisTitle: true, serAxisTitle: 'Series' },
		{ serAxisTitleFontSize: 20 }
	),
	chartOn(
		'addChart.serAxisTitleRotate',
		'bar3D',
		{ showSerAxisTitle: true, serAxisTitle: 'Series' },
		{ serAxisTitleRotate: 45 }
	),
	chartWith(
		'addChart.stockStyle:ohlc',
		'stock',
		[{ name: 'Open', labels: ['Mon', 'Tue', 'Wed'], values: [9, 11, 12] }, ...STOCK],
		{},
		{ stockStyle: 'ohlc' },
		['addChart.stockStyle']
	),
	chartOn('addChart.surface3D', 'surface', {}, { surface3D: false }),
	chartOn('addChart.surfaceWireframe', 'surface', {}, { surfaceWireframe: true }),
	chartWith('addChart.subtotals', 'waterfall', [CHART_SERIES[0] as OptsChartData], {}, { subtotals: [2] }),
	chartWith('addChart.binning.intervalClosed:l', 'histogram', OBSERVATIONS, {}, { binning: { intervalClosed: 'l' } }, [
		'addChart.binning.intervalClosed',
	]),
	chartWith(
		'addChart.statistics.quartileMethod:inclusive',
		'boxWhisker',
		OBSERVATIONS,
		{},
		{ statistics: { quartileMethod: 'inclusive' } },
		['addChart.statistics.quartileMethod']
	),
	chartWith('addChart.statistics.meanLine', 'boxWhisker', OBSERVATIONS, {}, { statistics: { meanLine: true } }),
	chartWith('addChart.statistics.meanMarker', 'boxWhisker', OBSERVATIONS, {}, { statistics: { meanMarker: false } }),
	chartWith('addChart.statistics.outliers', 'boxWhisker', OBSERVATIONS, {}, { statistics: { outliers: false } }),
	chartWith('addChart.statistics.nonoutliers', 'boxWhisker', OBSERVATIONS, {}, { statistics: { nonoutliers: true } }),
	chartWith(
		'addChart.geography.cultureLanguage',
		'regionMap',
		REGIONS,
		{},
		{ geography: { cultureLanguage: 'fr-FR' } }
	),
	chartWith('addChart.geography.cultureRegion', 'regionMap', REGIONS, {}, { geography: { cultureRegion: 'FR' } }),
	chart('addChart.x', { x: 2 }),
	chart('addChart.y', { y: 0.5 }),
	chart('addChart.w', { w: 6 }),
	chart('addChart.h', { h: 3 }),
	{
		...chart('addChart.objectLock.noAdjustHandles', { objectLock: { noAdjustHandles: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'CT_GraphicalObjectFrameLocking has no noAdjustHandles attribute, so a chart frame cannot carry it; ObjectLockProps is shared with shapes, where it applies',
		},
	},
	{
		...chart('addChart.objectLock.noChangeArrowheads', { objectLock: { noChangeArrowheads: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'CT_GraphicalObjectFrameLocking has no noChangeArrowheads attribute, so a chart frame cannot carry it; ObjectLockProps is shared with shapes, where it applies',
		},
	},
	chart('addChart.objectLock.noChangeAspect', { objectLock: { noChangeAspect: true } }),
	{
		...chart('addChart.objectLock.noChangeShapeType', { objectLock: { noChangeShapeType: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'CT_GraphicalObjectFrameLocking has no noChangeShapeType attribute, so a chart frame cannot carry it; ObjectLockProps is shared with shapes, where it applies',
		},
	},
	{
		...chart('addChart.objectLock.noCrop', { objectLock: { noCrop: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'CT_GraphicalObjectFrameLocking has no noCrop attribute, so a chart frame cannot carry it; ObjectLockProps is shared with shapes, where it applies',
		},
	},
	chart('addChart.objectLock.noDrilldown', { objectLock: { noDrilldown: true } }),
	{
		...chart('addChart.objectLock.noEditPoints', { objectLock: { noEditPoints: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'CT_GraphicalObjectFrameLocking has no noEditPoints attribute, so a chart frame cannot carry it; ObjectLockProps is shared with shapes, where it applies',
		},
	},
	chart('addChart.objectLock.noGrp', { objectLock: { noGrp: true } }),
	chart('addChart.objectLock.noMove', { objectLock: { noMove: true } }),
	chart('addChart.objectLock.noResize', { objectLock: { noResize: true } }),
	{
		...chart('addChart.objectLock.noRot', { objectLock: { noRot: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'CT_GraphicalObjectFrameLocking has no noRot attribute, so a chart frame cannot carry it; ObjectLockProps is shared with shapes, where it applies',
		},
	},
	chart('addChart.objectLock.noSelect', { objectLock: { noSelect: true } }),
	{
		...chart('addChart.objectLock.noTextEdit', { objectLock: { noTextEdit: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'CT_GraphicalObjectFrameLocking has no noTextEdit attribute, so a chart frame cannot carry it; ObjectLockProps is shared with shapes, where it applies',
		},
	},
	{
		...chartOn('addChart.align', 'bar', { showTitle: true, title: 'Probe title' }, { align: 'right' }),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn('addChart.bold', 'bar', { showTitle: true, title: 'Probe title' }, { bold: true }),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn('addChart.breakLine', 'bar', { showTitle: true, title: 'Probe title' }, { breakLine: true }),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn('addChart.caps', 'bar', { showTitle: true, title: 'Probe title' }, { caps: 'all' }),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn('addChart.color', 'bar', { showTitle: true, title: 'Probe title' }, { color: '9B1B30' }),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn('addChart.fontFace', 'bar', { showTitle: true, title: 'Probe title' }, { fontFace: 'Georgia' }),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn('addChart.fontFaceEA', 'bar', { showTitle: true, title: 'Probe title' }, { fontFaceEA: 'MS Mincho' }),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn('addChart.fontSize', 'bar', { showTitle: true, title: 'Probe title' }, { fontSize: 30 }),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn('addChart.highlight', 'bar', { showTitle: true, title: 'Probe title' }, { highlight: 'F0C808' }),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn('addChart.italic', 'bar', { showTitle: true, title: 'Probe title' }, { italic: true }),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn('addChart.paraIndent', 'bar', { showTitle: true, title: 'Probe title' }, { paraIndent: 12 }),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn('addChart.paraMarginLeft', 'bar', { showTitle: true, title: 'Probe title' }, { paraMarginLeft: 18 }),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn('addChart.softBreakBefore', 'bar', { showTitle: true, title: 'Probe title' }, { softBreakBefore: true }),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn('addChart.textDirection', 'bar', { showTitle: true, title: 'Probe title' }, { textDirection: 'vert' }),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn('addChart.textWarp', 'bar', { showTitle: true, title: 'Probe title' }, { textWarp: 'textArchUp' }),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn('addChart.transparency', 'bar', { showTitle: true, title: 'Probe title' }, { transparency: 40 }),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn('addChart.valign', 'bar', { showTitle: true, title: 'Probe title' }, { valign: 'bottom' }),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn('addChart.bullet', 'bar', { showTitle: true, title: 'Probe title' }, { bullet: true }),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn(
			'addChart.bullet.characterCode',
			'bar',
			{ showTitle: true, title: 'Probe title', bullet: true },
			{ bullet: { characterCode: '25BA' } }
		),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn(
			'addChart.bullet.color',
			'bar',
			{ showTitle: true, title: 'Probe title', bullet: true },
			{ bullet: { color: '9B1B30' } }
		),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn(
			'addChart.bullet.fontFace',
			'bar',
			{ showTitle: true, title: 'Probe title', bullet: true },
			{ bullet: { fontFace: 'Wingdings' } }
		),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn(
			'addChart.bullet.indent',
			'bar',
			{ showTitle: true, title: 'Probe title', bullet: true },
			{ bullet: { indent: 10 } }
		),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn(
			'addChart.bullet.size',
			'bar',
			{ showTitle: true, title: 'Probe title', bullet: true },
			{ bullet: { size: 80 } }
		),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn(
			'addChart.bullet.image.data',
			'bar',
			{ showTitle: true, title: 'Probe title', bullet: true },
			{ bullet: { image: { data: PNG } } }
		),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn(
			'addChart.bullet.type:number',
			'bar',
			{ showTitle: true, title: 'Probe title' },
			{ bullet: { type: 'number' } },
			['addChart.bullet.type']
		),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn(
			'addChart.bullet.numberStartAt',
			'bar',
			{ showTitle: true, title: 'Probe title', bullet: { type: 'number' } },
			{ bullet: { type: 'number', numberStartAt: 5 } }
		),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn(
			'addChart.bullet.numberType',
			'bar',
			{ showTitle: true, title: 'Probe title', bullet: { type: 'number' } },
			{ bullet: { type: 'number', numberType: 'romanLcParenR' } }
		),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn(
			'addChart.underline',
			'bar',
			{ showTitle: true, title: 'Probe title' },
			{ underline: { style: 'dbl', color: '9B1B30' } },
			['addChart.underline.style', 'addChart.underline.color']
		),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn('addChart.underline:true', 'bar', { showTitle: true, title: 'Probe title' }, { underline: true }),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	{
		...chartOn(
			'addChart.tabStops',
			'bar',
			{ showTitle: true, title: 'Probe title' },
			{ tabStops: [{ position: 1, alignment: 'r' }] },
			['addChart.tabStops.position', 'addChart.tabStops.alignment']
		),
		accepted: {
			kind: 'inert',
			reason:
				'a chart styles its text through its own options (titleBold, legendFontSize, dataLabelColor, ...) and reads no TextBaseProps key; the writer warns that it has no effect (#138, answered by #162)',
		},
	},
	chartOn(
		'addChart.chartArea.fill.gradient.center',
		'bar',
		{
			chartArea: {
				fill: {
					type: 'gradient',
					gradient: {
						kind: 'radial',
						stops: [
							{ position: 0, color: '9B1B30' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			},
		},
		{
			chartArea: {
				fill: {
					type: 'gradient',
					gradient: {
						kind: 'radial',
						center: { x: 20, y: 80 },
						stops: [
							{ position: 0, color: '9B1B30' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			},
		},
		['addChart.chartArea.fill.gradient.center.x', 'addChart.chartArea.fill.gradient.center.y']
	),
	chart('addChart.chartArea.fill.image.data', { chartArea: { fill: { type: 'image', image: { data: PNG } } } }),
	chartOn(
		'addChart.chartArea.fill.image.crop',
		'bar',
		{ chartArea: { fill: { type: 'image', image: { data: PNG } } } },
		{ chartArea: { fill: { type: 'image', image: { data: PNG, crop: { l: 10, t: 20, r: 30, b: 5 } } } } },
		[
			'addChart.chartArea.fill.image.crop.l',
			'addChart.chartArea.fill.image.crop.t',
			'addChart.chartArea.fill.image.crop.r',
			'addChart.chartArea.fill.image.crop.b',
		]
	),
	chartOn(
		'addChart.plotArea.fill.gradient.center',
		'bar',
		{
			plotArea: {
				fill: {
					type: 'gradient',
					gradient: {
						kind: 'radial',
						stops: [
							{ position: 0, color: '9B1B30' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			},
		},
		{
			plotArea: {
				fill: {
					type: 'gradient',
					gradient: {
						kind: 'radial',
						center: { x: 20, y: 80 },
						stops: [
							{ position: 0, color: '9B1B30' },
							{ position: 100, color: '250F6B' },
						],
					},
				},
			},
		},
		['addChart.plotArea.fill.gradient.center.x', 'addChart.plotArea.fill.gradient.center.y']
	),
	chart('addChart.plotArea.fill.image.data', { plotArea: { fill: { type: 'image', image: { data: PNG } } } }),
	chartOn(
		'addChart.plotArea.fill.image.crop',
		'bar',
		{ plotArea: { fill: { type: 'image', image: { data: PNG } } } },
		{ plotArea: { fill: { type: 'image', image: { data: PNG, crop: { l: 10, t: 20, r: 30, b: 5 } } } } },
		[
			'addChart.plotArea.fill.image.crop.l',
			'addChart.plotArea.fill.image.crop.t',
			'addChart.plotArea.fill.image.crop.r',
			'addChart.plotArea.fill.image.crop.b',
		]
	),
	chartOn(
		'addChart.catAxes.catAxisHidden',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisHidden: true }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisCrossesAt',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisCrossesAt: 2 }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisLabelColor',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisLabelColor: '9B1B30' }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisLabelFontBold',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisLabelFontBold: true }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisLabelFontFace',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisLabelFontFace: 'Georgia' }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisLabelFontItalic',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisLabelFontItalic: true }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisLabelFontSize',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisLabelFontSize: 18 }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisLabelFrequency',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisLabelFrequency: 2 }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisLabelRotate',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisLabelRotate: 45 }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisLineColor',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisLineColor: '9B1B30' }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisLineShow',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisLineShow: false }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisLineSize',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisLineSize: 3 }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisLabelPos:high',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisLabelPos: 'high' }], valAxes: [{}] },
		['addChart.catAxes.catAxisLabelPos']
	),
	chartOn(
		'addChart.catAxes.catAxisLineStyle:dash',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisLineStyle: 'dash' }], valAxes: [{}] },
		['addChart.catAxes.catAxisLineStyle']
	),
	chartOn(
		'addChart.catAxes.catAxisMajorTickMark:cross',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisMajorTickMark: 'cross' }], valAxes: [{}] },
		['addChart.catAxes.catAxisMajorTickMark']
	),
	chartOn(
		'addChart.catAxes.catAxisMinorTickMark:in',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisMinorTickMark: 'in' }], valAxes: [{}] },
		['addChart.catAxes.catAxisMinorTickMark']
	),
	chartOn(
		'addChart.catAxes.catAxisOrientation:maxMin',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisOrientation: 'maxMin' }], valAxes: [{}] },
		['addChart.catAxes.catAxisOrientation']
	),
	chartOn(
		'addChart.catAxes.catAxisLabelFormatCode',
		'scatter',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisLabelFormatCode: '0.00' }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisMajorUnit',
		'scatter',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisMajorUnit: 2 }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisMinorUnit',
		'scatter',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisMinorUnit: 0.5 }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisMaxVal',
		'scatter',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisMaxVal: 10 }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisMinVal',
		'scatter',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisMinVal: 1 }], valAxes: [{}] }
	),
	chartWith(
		'addChart.catAxes.catLabelFormatCode',
		'bar',
		DATED,
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catLabelFormatCode: 'yyyy-mm-dd' }], valAxes: [{}] }
	),
	chartWith(
		'addChart.catAxes.catAxisBaseTimeUnit:months',
		'bar',
		DATED,
		{ catAxes: [{ catLabelFormatCode: 'yyyy-mm-dd' }], valAxes: [{}] },
		{ catAxes: [{ catLabelFormatCode: 'yyyy-mm-dd', catAxisBaseTimeUnit: 'months' }], valAxes: [{}] },
		['addChart.catAxes.catAxisBaseTimeUnit']
	),
	chartWith(
		'addChart.catAxes.catAxisMajorTimeUnit:months',
		'bar',
		DATED,
		{ catAxes: [{ catLabelFormatCode: 'yyyy-mm-dd' }], valAxes: [{}] },
		{ catAxes: [{ catLabelFormatCode: 'yyyy-mm-dd', catAxisMajorTimeUnit: 'months' }], valAxes: [{}] },
		['addChart.catAxes.catAxisMajorTimeUnit']
	),
	chartWith(
		'addChart.catAxes.catAxisMinorTimeUnit:days',
		'bar',
		DATED,
		{ catAxes: [{ catLabelFormatCode: 'yyyy-mm-dd' }], valAxes: [{}] },
		{ catAxes: [{ catLabelFormatCode: 'yyyy-mm-dd', catAxisMinorTimeUnit: 'days' }], valAxes: [{}] },
		['addChart.catAxes.catAxisMinorTimeUnit']
	),
	chartWith(
		'addChart.catAxes.catAxisMultiLevelLabels',
		'bar',
		LEVELS,
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisMultiLevelLabels: true }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.showCatAxisTitle',
		'bar',
		{ catAxes: [{ catAxisTitle: 'Category' }], valAxes: [{}] },
		{ catAxes: [{ catAxisTitle: 'Category', showCatAxisTitle: true }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisTitle',
		'bar',
		{ catAxes: [{ showCatAxisTitle: true }], valAxes: [{}] },
		{ catAxes: [{ showCatAxisTitle: true, catAxisTitle: 'Category' }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisTitleColor',
		'bar',
		{ catAxes: [{ showCatAxisTitle: true, catAxisTitle: 'Category' }], valAxes: [{}] },
		{ catAxes: [{ showCatAxisTitle: true, catAxisTitle: 'Category', catAxisTitleColor: '9B1B30' }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisTitleFontFace',
		'bar',
		{ catAxes: [{ showCatAxisTitle: true, catAxisTitle: 'Category' }], valAxes: [{}] },
		{ catAxes: [{ showCatAxisTitle: true, catAxisTitle: 'Category', catAxisTitleFontFace: 'Georgia' }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisTitleFontSize',
		'bar',
		{ catAxes: [{ showCatAxisTitle: true, catAxisTitle: 'Category' }], valAxes: [{}] },
		{ catAxes: [{ showCatAxisTitle: true, catAxisTitle: 'Category', catAxisTitleFontSize: 20 }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisTitleRotate',
		'bar',
		{ catAxes: [{ showCatAxisTitle: true, catAxisTitle: 'Category' }], valAxes: [{}] },
		{ catAxes: [{ showCatAxisTitle: true, catAxisTitle: 'Category', catAxisTitleRotate: 45 }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisLine.cap',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisLine: { cap: 'round' } }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisLine.color',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisLine: { color: '9B1B30' } }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisLine.dashType',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisLine: { dashType: 'lgDashDot' } }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisLine.width',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisLine: { width: 3 } }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catAxisLine.type:dash',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catAxisLine: { type: 'dash' } }], valAxes: [{}] },
		['addChart.catAxes.catAxisLine.type']
	),
	chartOn(
		'addChart.catAxes.catAxisLine.transparency',
		'bar',
		{ catAxes: [{ catAxisLine: { color: '9B1B30' } }], valAxes: [{}] },
		{ catAxes: [{ catAxisLine: { color: '9B1B30', transparency: 40 } }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catGridLine.cap',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catGridLine: { cap: 'round' } }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catGridLine.color',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catGridLine: { color: '9B1B30' } }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catGridLine.dashType',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catGridLine: { dashType: 'lgDashDot' } }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catGridLine.width',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catGridLine: { width: 3 } }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catGridLine.size',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catGridLine: { size: 3 } }], valAxes: [{}] }
	),
	chartOn(
		'addChart.catAxes.catGridLine.type:dash',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catGridLine: { type: 'dash' } }], valAxes: [{}] },
		['addChart.catAxes.catGridLine.type']
	),
	chartOn(
		'addChart.catAxes.catGridLine.style:dash',
		'bar',
		{ catAxes: [{}], valAxes: [{}] },
		{ catAxes: [{ catGridLine: { style: 'dash' } }], valAxes: [{}] },
		['addChart.catAxes.catGridLine.style']
	),
	chartOn(
		'addChart.catAxes.catGridLine.transparency',
		'bar',
		{ catAxes: [{ catGridLine: { color: '9B1B30' } }], valAxes: [{}] },
		{ catAxes: [{ catGridLine: { color: '9B1B30', transparency: 40 } }], valAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisHidden',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisHidden: true }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisCrossesAt',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisCrossesAt: 2 }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisDisplayUnitLabel',
		'bar',
		{ valAxes: [{ valAxisDisplayUnit: 'thousands' }], catAxes: [{}] },
		{ valAxes: [{ valAxisDisplayUnit: 'thousands', valAxisDisplayUnitLabel: true }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisLabelColor',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisLabelColor: '9B1B30' }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisLabelFontBold',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisLabelFontBold: true }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisLabelFontFace',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisLabelFontFace: 'Georgia' }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisLabelFontItalic',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisLabelFontItalic: true }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisLabelFontSize',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisLabelFontSize: 18 }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisLabelFormatCode',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisLabelFormatCode: '0.00' }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisLabelRotate',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisLabelRotate: 45 }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisLineColor',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisLineColor: '9B1B30' }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisLineShow',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisLineShow: false }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisLineSize',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisLineSize: 3 }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisLogScaleBase',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisLogScaleBase: 10 }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisMajorUnit',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisMajorUnit: 2 }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisMaxVal',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisMaxVal: 12 }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisMinorUnit',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisMinorUnit: 0.5 }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisMinVal',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisMinVal: 1 }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valLabelFormatCode',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valLabelFormatCode: '0.0%' }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisCrossBetween:midCat',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisCrossBetween: 'midCat' }], catAxes: [{}] },
		['addChart.valAxes.valAxisCrossBetween']
	),
	chartOn(
		'addChart.valAxes.valAxisDisplayUnit:thousands',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisDisplayUnit: 'thousands' }], catAxes: [{}] },
		['addChart.valAxes.valAxisDisplayUnit']
	),
	chartOn(
		'addChart.valAxes.valAxisLabelPos:high',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisLabelPos: 'high' }], catAxes: [{}] },
		['addChart.valAxes.valAxisLabelPos']
	),
	chartOn(
		'addChart.valAxes.valAxisLineStyle:dash',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisLineStyle: 'dash' }], catAxes: [{}] },
		['addChart.valAxes.valAxisLineStyle']
	),
	chartOn(
		'addChart.valAxes.valAxisMajorTickMark:cross',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisMajorTickMark: 'cross' }], catAxes: [{}] },
		['addChart.valAxes.valAxisMajorTickMark']
	),
	chartOn(
		'addChart.valAxes.valAxisMinorTickMark:in',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisMinorTickMark: 'in' }], catAxes: [{}] },
		['addChart.valAxes.valAxisMinorTickMark']
	),
	chartOn(
		'addChart.valAxes.valAxisOrientation:maxMin',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisOrientation: 'maxMin' }], catAxes: [{}] },
		['addChart.valAxes.valAxisOrientation']
	),
	chartOn(
		'addChart.valAxes.showValAxisTitle',
		'bar',
		{ valAxes: [{ valAxisTitle: 'Value' }], catAxes: [{}] },
		{ valAxes: [{ valAxisTitle: 'Value', showValAxisTitle: true }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisTitle',
		'bar',
		{ valAxes: [{ showValAxisTitle: true }], catAxes: [{}] },
		{ valAxes: [{ showValAxisTitle: true, valAxisTitle: 'Value' }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisTitleColor',
		'bar',
		{ valAxes: [{ showValAxisTitle: true, valAxisTitle: 'Value' }], catAxes: [{}] },
		{ valAxes: [{ showValAxisTitle: true, valAxisTitle: 'Value', valAxisTitleColor: '9B1B30' }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisTitleFontFace',
		'bar',
		{ valAxes: [{ showValAxisTitle: true, valAxisTitle: 'Value' }], catAxes: [{}] },
		{ valAxes: [{ showValAxisTitle: true, valAxisTitle: 'Value', valAxisTitleFontFace: 'Georgia' }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisTitleFontSize',
		'bar',
		{ valAxes: [{ showValAxisTitle: true, valAxisTitle: 'Value' }], catAxes: [{}] },
		{ valAxes: [{ showValAxisTitle: true, valAxisTitle: 'Value', valAxisTitleFontSize: 20 }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisTitleRotate',
		'bar',
		{ valAxes: [{ showValAxisTitle: true, valAxisTitle: 'Value' }], catAxes: [{}] },
		{ valAxes: [{ showValAxisTitle: true, valAxisTitle: 'Value', valAxisTitleRotate: 45 }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisLine.cap',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisLine: { cap: 'round' } }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisLine.color',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisLine: { color: '9B1B30' } }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisLine.dashType',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisLine: { dashType: 'lgDashDot' } }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisLine.width',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisLine: { width: 3 } }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valAxisLine.type:dash',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valAxisLine: { type: 'dash' } }], catAxes: [{}] },
		['addChart.valAxes.valAxisLine.type']
	),
	chartOn(
		'addChart.valAxes.valAxisLine.transparency',
		'bar',
		{ valAxes: [{ valAxisLine: { color: '9B1B30' } }], catAxes: [{}] },
		{ valAxes: [{ valAxisLine: { color: '9B1B30', transparency: 40 } }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valGridLine.cap',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valGridLine: { cap: 'round' } }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valGridLine.color',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valGridLine: { color: '9B1B30' } }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valGridLine.dashType',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valGridLine: { dashType: 'lgDashDot' } }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valGridLine.width',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valGridLine: { width: 3 } }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valGridLine.size',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valGridLine: { size: 3 } }], catAxes: [{}] }
	),
	chartOn(
		'addChart.valAxes.valGridLine.type:dash',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valGridLine: { type: 'dash' } }], catAxes: [{}] },
		['addChart.valAxes.valGridLine.type']
	),
	chartOn(
		'addChart.valAxes.valGridLine.style:dash',
		'bar',
		{ valAxes: [{}], catAxes: [{}] },
		{ valAxes: [{ valGridLine: { style: 'dash' } }], catAxes: [{}] },
		['addChart.valAxes.valGridLine.style']
	),
	chartOn(
		'addChart.valAxes.valGridLine.transparency',
		'bar',
		{ valAxes: [{ valGridLine: { color: '9B1B30' } }], catAxes: [{}] },
		{ valAxes: [{ valGridLine: { color: '9B1B30', transparency: 40 } }], catAxes: [{}] }
	),
	chartCombo('addChart.secondaryCatAxis', { secondaryCatAxis: true }),
	chartCombo('addChart.catAxes.secondaryCatAxis', { secondaryCatAxis: true }),
	chartCombo('addChart.secondaryValAxis', { secondaryValAxis: true }),
	chartCombo('addChart.valAxes.secondaryValAxis', { secondaryValAxis: true }),
	{
		...chartOn(
			'addChart.catAxes.catAxes',
			'bar',
			{ catAxes: [{}], valAxes: [{}] },
			{ catAxes: [{ catAxes: [{ showCatAxisTitle: true, catAxisTitle: 'Nested' }] }], valAxes: [{}] }
		),
		accepted: {
			kind: 'inert',
			reason:
				'an element of the axis array is read as one axis; an axis array nested inside it is read by nothing, and is only reachable because the element type is the same props interface the array hangs off',
		},
	},
	{
		...chartOn(
			'addChart.valAxes.valAxes',
			'bar',
			{ catAxes: [{}], valAxes: [{}] },
			{ valAxes: [{ valAxes: [{ showValAxisTitle: true, valAxisTitle: 'Nested' }] }], catAxes: [{}] }
		),
		accepted: {
			kind: 'inert',
			reason:
				'an element of the axis array is read as one axis; an axis array nested inside it is read by nothing, and is only reachable because the element type is the same props interface the array hangs off',
		},
	},
	{
		id: 'addChart.type:line',
		covers: ['addChart.type'],
		author: (pptx: TsPptx, on: boolean) => {
			pptx.addSlide().addChart(CHART_SERIES, { ...BOX, h: 4, objectName: 'probe', type: on ? 'line' : 'bar' })
		},
	},
	chartPlaceholder('addChart.placeholder', undefined, 'chart'),
	chartPlaceholder('addChart.placeholder.name', undefined, { name: 'chart' }),
	chartPlaceholder('addChart.placeholder.type', undefined, { type: 'chart' }),
	chartPlaceholder('addChart.placeholder.idx', { type: 'chart' }, { type: 'chart', idx: 13 }),
	{
		...chart('addChart.objectLock.noUngrp', { objectLock: { noUngrp: true } }),
		accepted: {
			kind: 'inert',
			reason:
				'CT_GraphicalObjectFrameLocking has no noUngrp attribute, so a chart frame cannot carry it; ObjectLockProps is shared with groups, where it applies',
		},
	},
]

/**
 * Options that change the XML and are never painted, in any context: the render
 * frontier's `invisible`, by category. A probe whose id matches takes the reason.
 *
 * A category belongs here only when *no* slide would draw it. One that would be
 * painted in context, and the probe gave it none, is a probe to fix, not an entry
 * to add. A rule cannot hide a render regression: a probe it claims that renders
 * fails on the stale claim, and one that stops rendering falls in rank.
 */
const INVISIBLE: [RegExp, string][] = [
	[/\.objectLock\./, 'an editing lock (`a:*Locks`): restricts what PowerPoint lets a user change, and draws nothing'],
	[
		/^slide\.transition\./,
		'a slide transition: plays between slides in a slideshow, and draws nothing on a static slide',
	],
	[/^addAnimation\./, 'an animation: plays during a slideshow, and draws nothing on a static slide'],
	[/^addMedia\.loop/, 'media playback: changes how the clip plays, not the frame drawn for it'],
	[/\.lang(:|$)/, 'a proofing language (`a:rPr/@lang`): selects spelling and grammar rules, and draws nothing'],
	[
		/^add(Shape|Image)\.hyperlink\./,
		'a click action on the whole object (`p:cNvPr/a:hlinkClick`): fires in a slideshow, and draws nothing; a run hyperlink, which styles its text, is not this',
	],
	[
		/background\.gradient\.rotateWithShape$/,
		'whether a gradient turns with its shape (`a:gradFill/@rotWithShape`): a background has no rotation to follow, so it draws the same either way',
	],
	[
		/^addShape\.(connectionSites|adjustHandles)\b/,
		'a custom geometry editing affordance (`a:cxnLst` / `a:ahLst`): where connectors snap and which handles PowerPoint offers, neither of which is drawn',
	],
]

export const PROBES: Probe[] = AUTHORED.map((probe) => {
	const reason = INVISIBLE.find(([pattern]) => pattern.test(probe.id))?.[1]
	return reason ? { ...probe, invisible: reason } : probe
})

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
	'addChart.bullet.image.path':
		'the file-path spelling of the bullet picture source; the bullet it would write is exercised by addChart.bullet.image.data',
	'addChart.chartArea.fill.image.path':
		'the file-path spelling of the picture source; the picture it writes is exercised by addChart.chartArea.fill.image.data',
	'addChart.plotArea.fill.image.path':
		'the file-path spelling of the picture source; the picture it writes is exercised by addChart.plotArea.fill.image.data',
	'defineSlideMaster.objects.chart.type.data.sizes':
		'bubble sizes are read only by a bubble chart, and the writer refuses a bubble chart in the multi-type form by design (InvalidOptionError: PowerPoint refuses the deck), so this spelling cannot be authored (ts-pptx#140: https://github.com/shbernal/pptx-ts/issues/140 asks for the type to stop admitting it); the sizes it would write are exercised by defineSlideMaster.objects.chart.data.sizes',
	'addMedia.data':
		'the media source, which every addMedia probe sets on both sides; the media part it writes is exercised by addMedia.extn, which renames that part',
	'addMedia.path': 'the file-path spelling of the media source; the media part it writes is exercised by addMedia.extn',
	'addAnimation.shapeIndex':
		'selects the target by add order, the alternative to objectName; all it changes is p:spTgt@spid and p:bldP@spid, which the fact diff masks as ids, and the effect it targets is exercised by addAnimation.preset',
	'addTable.autoPageCharWeight':
		'an auto-paging knob that only weighs where autoPage splits a table, and writes no XML of its own',
	'addTable.autoPageLineWeight':
		'an auto-paging knob that only weighs where autoPage splits a table, and writes no XML of its own',
	'addTable.bullet.image.path':
		'the file-path spelling of the bullet picture source; the bullet it writes is exercised by addTable.bullet.image.data',
	'addTable.fill.image.path':
		'the file-path spelling of the picture source; the picture it writes is exercised by addTable.fill.image.data',
	'addTable.tableFill.image.path':
		'the file-path spelling of the picture source; the picture it writes is exercised by addTable.tableFill.image.data',
	'addTable.verbose': 'a development flag that logs the auto-paging trace to the console and writes no XML of its own',
}
