/**
 * Read model → IR text.
 *
 * The one place the "absent means inherited" rule earns its keep, and the one
 * place it is not enough on its own. A run in a layout placeholder often states
 * *nothing* — no size, no face, no colour — and a renderer cannot paint nothing.
 * So each run carries both halves:
 *
 * - `props` is what the run itself said. An absent field means inherited, and
 *   emit reads only this, so a layout's 44pt title never gets baked into a slide.
 * - `resolved` is what to paint, after the read model has walked the
 *   placeholder → layout → master → `defaultTextStyle` chain.
 *
 * Writing the resolved value into `props` would collapse the two and pass a
 * pixel diff while failing a structural one — the flattening trap in miniature.
 *
 * Frame-level properties (`anchor`, `wrap`, insets, autofit) *are* resolved to a
 * concrete value, and that is not the same mistake: they describe layout inside
 * the box, nothing emits from `TextBody`, and PowerPoint's own defaults for them
 * are fixed rather than inherited from a placeholder.
 */

import type {
	AutofitMode,
	BulletDetail,
	BulletStyle,
	Paragraph as ReadParagraph,
	Run as ReadRun,
	TextFrame,
} from '@shbernal/ts-pptx/read'
import type {
	Bullet,
	Color,
	Hyperlink,
	Paragraph,
	ParagraphProperties,
	ResolvedRunProperties,
	RunProperties,
	TextBody,
	TextRun,
} from '../ir/render'
import { type ImportScope, note } from './context'
import { colorOf } from './paint'

/** PowerPoint's `a:bodyPr` inset defaults: 0.1" horizontally, 0.05" vertically. */
const DEFAULT_INSETS = { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 } as const

const ALIGN: Record<string, ParagraphProperties['align']> = {
	l: 'left',
	ctr: 'center',
	r: 'right',
	just: 'justify',
}

const ANCHOR: Record<string, TextBody['anchor']> = { t: 'top', ctr: 'middle', b: 'bottom' }

const AUTOFIT: Record<AutofitMode, TextBody['autofit']> = {
	none: 'none',
	normAutofit: 'shrink',
	spAutoFit: 'resize',
}

const VERTICAL = new Set<string>(['vert', 'vert270', 'eaVert', 'wordArtVert'])

/** `ST_TextUnderlineType` has seventeen members; the write API expresses three. */
const UNDERLINE: Record<string, NonNullable<RunProperties['underline']>> = {
	none: 'none',
	sng: 'single',
	dbl: 'double',
}

const STRIKE: Record<string, NonNullable<RunProperties['strike']>> = {
	noStrike: 'none',
	sngStrike: 'single',
	dblStrike: 'double',
}

export function textBodyOf(frame: TextFrame, scope: ImportScope): TextBody {
	const body = frame.bodyProperties
	const anchor = frame.resolvedAnchor
	const insets = body?.insetsPt ?? {}

	const vertical = body?.vert ?? null
	let modeledVertical: TextBody['vertical'] | undefined
	if (vertical !== null && vertical !== 'horz') {
		if (VERTICAL.has(vertical)) modeledVertical = vertical as TextBody['vertical']
		else
			note(
				scope,
				'text.vert',
				'dropped',
				'unsupported',
				`a text direction of ${JSON.stringify(vertical)} has no modeled equivalent, so the text is laid out horizontally`
			)
	}

	return {
		paragraphs: frame.paragraphs.map((paragraph) => paragraphOf(paragraph, scope)),
		autofit: AUTOFIT[frame.autofit ?? 'none'],
		anchor: (anchor !== null ? ANCHOR[anchor] : undefined) ?? 'top',
		// `@wrap` is `square` (wrap) or `none`; unset means wrap.
		wrap: body?.wrap !== 'none',
		insetsPt: {
			left: insets.left ?? DEFAULT_INSETS.left,
			right: insets.right ?? DEFAULT_INSETS.right,
			top: insets.top ?? DEFAULT_INSETS.top,
			bottom: insets.bottom ?? DEFAULT_INSETS.bottom,
		},
		...(modeledVertical === undefined ? {} : { vertical: modeledVertical }),
	}
}

function paragraphOf(paragraph: ReadParagraph, scope: ImportScope): Paragraph {
	const align = paragraph.align
	let modeledAlign: ParagraphProperties['align']
	if (align !== null) {
		modeledAlign = ALIGN[align]
		if (modeledAlign === undefined)
			note(
				scope,
				'text.align',
				'dropped',
				'unwritable',
				`an alignment of ${JSON.stringify(align)} has no write-API equivalent, so the paragraph inherits its alignment instead`
			)
	}

	const props: ParagraphProperties = {
		level: paragraph.level,
		...(modeledAlign === undefined ? {} : { align: modeledAlign }),
		...bulletFields(paragraph.bulletDetail, scope),
		...(paragraph.marginLeftPt === null ? {} : { marginLeftPt: paragraph.marginLeftPt }),
		...(paragraph.indentPt === null ? {} : { indentPt: paragraph.indentPt }),
		...(paragraph.lineSpacing === null ? {} : { lineSpacing: paragraph.lineSpacing }),
		...(paragraph.spaceBeforePt === null ? {} : { spaceBeforePt: paragraph.spaceBeforePt }),
		...(paragraph.spaceAfterPt === null ? {} : { spaceAfterPt: paragraph.spaceAfterPt }),
	}

	return { props, runs: paragraph.runs.map((run) => runOf(run, scope)) }
}

/**
 * The bullet's own `a:buFont` / `a:buSzPct` / `a:buClr`, shared by the three arms
 * that have a glyph to style.
 *
 * `a:buSzPts` is a note rather than a field: the write API's `bullet.size` is a
 * percentage of the run size, so an absolute point size has no expression to come
 * back through. `text.bullet.sizePt` is upstream's own key for exactly this, so
 * this lane's note set and the script tier's describe the one loss the same way.
 */
function bulletStyleFields(
	style: BulletStyle,
	scope: ImportScope
): { font?: string; color?: Color; sizePct?: number } {
	if (style.sizePt !== null)
		note(
			scope,
			'text.bullet.sizePt',
			'dropped',
			'unwritable',
			`this bullet sets an absolute glyph size of ${style.sizePt}pt (a:buSzPts); bullet.size is a percentage of the run size, so the glyph follows the text size instead`
		)

	const color = colorOf(style.schemeColor, style.resolvedColor) ?? colorOfHex(style.color)
	return {
		...(style.font === null ? {} : { font: style.font }),
		...(color === undefined ? {} : { color }),
		...(style.sizePct === null ? {} : { sizePct: style.sizePct }),
	}
}

/** `a:buClr` with a literal `a:srgbClr` and no theme context to resolve it against. */
function colorOfHex(hex: string | null): Color | undefined {
	return hex === null ? undefined : { kind: 'srgb', hex }
}

/**
 * `Paragraph.bulletDetail` — a discriminated union as of ts-pptx 3.0.0, which
 * replaced the tagged string (`'none'` / `'char:•'` / `'autoNum:arabicPeriod'`)
 * this used to parse ({@link https://github.com/shbernal/ts-pptx/issues/3}). The
 * old accessor was ambiguous for a glyph that is itself a colon, and had no room
 * for `@startAt` or the bullet's own font/size/colour, all of which land here now.
 *
 * `null` still has to stay absent rather than becoming `{ kind: 'none' }`: a
 * paragraph inheriting the layout's bullet and one that explicitly suppresses its
 * bullet render differently.
 */
function bulletFields(bullet: BulletDetail | null, scope: ImportScope): { bullet?: Bullet } {
	if (bullet === null) return {}
	if (bullet.kind === 'none') return { bullet: { kind: 'none' } }

	const style = bulletStyleFields(bullet, scope)
	if (bullet.kind === 'char') return { bullet: { kind: 'character', char: bullet.char, ...style } }
	if (bullet.kind === 'autoNum') {
		return {
			bullet: {
				kind: 'number',
				scheme: bullet.scheme,
				...(bullet.startAt === null ? {} : { startAt: bullet.startAt }),
				...style,
			},
		}
	}

	// A picture bullet (`a:buBlip`). Modeled when its image resolves, because the
	// renderer can paint it; the emit-side loss is upstream's `text.bullet.picture`
	// note, not this lane's to restate.
	const asset = bullet.imagePartName === null ? null : scope.assets.refFor(bullet.imagePartName)
	if (asset === null) {
		note(
			scope,
			'text.bullet.picture',
			'dropped',
			'unread',
			'this paragraph uses an image as its bullet glyph (a:buBlip) that resolves to no media part, so there are no bytes to draw and the bullet is inherited instead'
		)
		return {}
	}
	return { bullet: { kind: 'picture', asset, ...style } }
}

function runOf(run: ReadRun, scope: ImportScope): TextRun {
	return { text: run.text, props: runPropsOf(run, scope), resolved: resolvedRunPropsOf(run) }
}

function runPropsOf(run: ReadRun, scope: ImportScope): RunProperties {
	// The own-colour accessors, never the resolved one: `resolvedColor` answers for
	// a run that stated nothing, and copying that answer here is what would bake a
	// placeholder's inherited colour into the slide.
	const statesColor = run.color !== null || run.schemeColor !== null
	const color = statesColor
		? (colorOf(run.schemeColor, run.resolvedColor) ??
			(run.color === null ? undefined : ({ kind: 'srgb', hex: run.color } as const)))
		: undefined

	return {
		...(run.fontName === null ? {} : { fontFace: run.fontName }),
		...(run.fontSizePt === null ? {} : { sizePt: run.fontSizePt }),
		...(run.bold === null ? {} : { bold: run.bold }),
		...(run.italic === null ? {} : { italic: run.italic }),
		...enumField(run.underline, UNDERLINE, 'underline', scope, 'text.underline', 'underline style'),
		...enumField(run.strike, STRIKE, 'strike', scope, 'text.strike', 'strikethrough style'),
		...(color === undefined ? {} : { color }),
		...(run.charSpacingPt === null ? {} : { spacingPt: run.charSpacingPt }),
		...(run.baselinePct === null ? {} : { baselinePct: run.baselinePct }),
		...hyperlinkFields(run, scope),
	}
}

/**
 * Map an OOXML token onto the subset this IR models, or file a note.
 *
 * `text.underline` and `text.strike` are keys upstream does not have, because its
 * own notes never reach either construct. Coining them is the intended move — the
 * rule is not to invent a *synonym* for an existing key, and there is no existing
 * key for these. The alternative is rounding `wavyHeavy` to `single`, which is an
 * approximation with no way back.
 */
function enumField<K extends string, V extends string>(
	value: string | null,
	table: Record<string, V>,
	field: K,
	scope: ImportScope,
	construct: string,
	label: string
): Partial<Record<K, V>> {
	if (value === null) return {}
	const mapped = table[value]
	if (mapped !== undefined) return { [field]: mapped } as Partial<Record<K, V>>
	note(
		scope,
		construct,
		'dropped',
		'unwritable',
		`a ${label} of ${JSON.stringify(value)} is outside the three the write API expresses, so the run carries none`
	)
	return {}
}

function hyperlinkFields(run: ReadRun, scope: ImportScope): { hyperlink?: Hyperlink } {
	const link = run.hyperlink
	if (link === null) return {}
	const targetSlide = link.targetPartName === null ? undefined : scope.slideNumberByPart.get(link.targetPartName)
	const hyperlink: Hyperlink = {
		url: link.url,
		...(targetSlide === undefined ? {} : { targetSlide }),
		...(link.tooltip === null ? {} : { tooltip: link.tooltip }),
	}
	return { hyperlink }
}

function resolvedRunPropsOf(run: ReadRun): ResolvedRunProperties {
	const resolved = run.resolvedColor
	return {
		...(run.resolvedFontFace === null ? {} : { fontFace: run.resolvedFontFace }),
		...(run.resolvedSizePt === null ? {} : { sizePt: run.resolvedSizePt }),
		...(run.resolvedBold === null ? {} : { bold: run.resolvedBold }),
		// Painted, so the token is deliberately gone: `props.color` is where the
		// theme reference survives, and this is the pixel it currently resolves to.
		...(resolved === null
			? {}
			: {
					color: {
						kind: 'srgb' as const,
						hex: resolved.effectiveHex,
						...(resolved.alpha === undefined ? {} : { alpha: resolved.alpha }),
					},
				}),
	}
}
