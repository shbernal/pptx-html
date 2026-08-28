/**
 * {@link TextBody} → HTML inside a `<foreignObject>`.
 *
 * ## Why HTML and not SVG `<text>`
 *
 * SVG text positions glyphs exactly and wraps nothing: every line break would
 * have to be computed here, which means measuring text, which means the same IR
 * renders differently on two machines. The renderer is allowed to be lossy about
 * the picture but the charter forbids measurement-derived decisions from leaking
 * back, and a bespoke line breaker is a large project whose output would still
 * not match PowerPoint's. The browser already has a line breaker; `foreignObject`
 * is how an SVG borrows it.
 *
 * The cost is real and bounded: `foreignObject` is unevenly supported outside
 * browsers. The render target *is* a browser, so this is a cost in a direction
 * nothing in scope travels.
 *
 * ## Units
 *
 * Every length below is **points written as `px`** — a 44pt run is
 * `font-size: 44px`. The frame is placed inside a group scaled by EMU-per-point
 * (see `textFrame` in `node.ts`), so one CSS pixel here is one point of slide
 * space rather than one EMU.
 *
 * That group is not decoration. The slide's `viewBox` is in EMU and inside a
 * `foreignObject` one CSS pixel is one user unit, so writing these lengths in EMU
 * — which this file used to do — asks for `font-size: 558800px`, and browsers cap
 * `font-size` (Chrome at 10000px). Every run above roughly a point rendered at the
 * same hairline size, so the picture came out blank while the island beside it was
 * exact. Nothing about the round trip noticed, which is the separation working;
 * it is still not a preview anybody can read.
 *
 * ## `props.X ?? resolved.X`, never one alone
 *
 * A layout placeholder's run states nothing: `props` is `{}` and the real 44pt
 * lives in `resolved`. Painting from `props` alone renders every placeholder
 * title at the browser default; painting from `resolved` alone throws away a
 * user's edit, because `resolved` is derived at import and an edit only ever
 * touches `props`. The fallback chain is the only reading that keeps both.
 */

import type { Bullet, Paragraph, ParagraphProperties, TextBody, TextRun } from '../ir/render'
import { editableParaProps, editableRunProps } from '../ir/surface'
import { cssColor, escapeAttr } from './paint'

/** Text-node escaping. `&` first, or it would double-escape the entities below. */
export function escapeText(value: string): string {
	return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
}

const ANCHOR: Record<TextBody['anchor'], string> = {
	top: 'flex-start',
	middle: 'center',
	bottom: 'flex-end',
}

/** Three decimals, finer than any unit here resolves. */
function tidy(value: number): number {
	return Math.round(value * 1000) / 1000
}

/** Points, tidied — this file's `px` is the frame's own point-scaled unit. */
function px(points: number): number {
	return tidy(points)
}

/**
 * What one line of "single" spacing is worth, as a multiple of the font size.
 *
 * Needed only to reduce a line spacing the paragraph never stated: a percentage
 * has a base to subtract from, an unstated spacing does not, and CSS has no way
 * to say "`normal`, less a fifth". Both PowerPoint's single spacing and a
 * browser's `line-height: normal` come from the font's own ascent, descent and
 * line gap, which for the faces these decks use lands near 1.2 — Calibri reports
 * 1.22. A stand-in, so the frames that use it say so.
 */
const SINGLE_LINE_HEIGHT = 1.2

/**
 * The size a run is painted at: what it states, or what it inherits, times the
 * frame's baked shrink.
 *
 * `fontScale` is a number PowerPoint wrote into the file, not one measured here,
 * so applying it is reading the deck rather than re-deriving its layout — the
 * distinction the no-measurement rule turns on.
 */
function scaled(sizePt: number, fontScalePct: number | undefined): number {
	return fontScalePct === undefined ? sizePt : (sizePt * fontScalePct) / 100
}

function runStyle(run: TextRun, fontScalePct: number | undefined): string {
	const props = run.props
	const resolved = run.resolved
	const style: string[] = []

	const face = props.fontFace ?? resolved.fontFace
	if (face !== undefined) style.push(`font-family:${JSON.stringify(face)},sans-serif`)

	const size = props.sizePt ?? resolved.sizePt
	if (size !== undefined) style.push(`font-size:${px(scaled(size, fontScalePct))}px`)

	const bold = props.bold ?? resolved.bold
	if (bold !== undefined) style.push(`font-weight:${bold ? 700 : 400}`)

	// Workaround for ts-pptx#27 - https://github.com/shbernal/ts-pptx/issues/27
	//
	// The one property on this list painted from `props` alone. Every neighbour is
	// `props.X ?? resolved.X` because the read model has `resolvedBold`,
	// `resolvedSizePt`, `resolvedFontFace` and `resolvedColor`; there is no
	// `Run.resolvedItalic`, so `ResolvedRunProperties` has no `italic` to fall back
	// to. A run that inherits italic from a layout placeholder's `a:defRPr i="1"`
	// is painted upright. Paint-only: `props` is what emit reads, so the deck still
	// round-trips.
	//
	// Remove when `Run.resolvedItalic` exists. Then add `italic` to
	// `ResolvedRunProperties` (`src/ir/render.ts`), map it in `resolvedRunPropsOf`
	// (`src/import/text.ts`), and write `props.italic ?? resolved.italic` here.
	if (props.italic !== undefined) style.push(`font-style:${props.italic ? 'italic' : 'normal'}`)

	const color = props.color ?? resolved.color
	if (color !== undefined) style.push(`color:${cssColor(color)}`)

	// `underline` and `strike` are both `none`/`single`/`double`; CSS spells the
	// last one as a separate style rather than a second line.
	const decoration: string[] = []
	if (props.underline !== undefined && props.underline !== 'none') decoration.push('underline')
	if (props.strike !== undefined && props.strike !== 'none') decoration.push('line-through')
	if (decoration.length > 0) {
		style.push(`text-decoration:${decoration.join(' ')}`)
		if (props.underline === 'double' || props.strike === 'double') style.push('text-decoration-style:double')
	}

	if (props.spacingPt !== undefined) style.push(`letter-spacing:${px(props.spacingPt)}px`)
	if (props.baselinePct !== undefined && props.baselinePct !== 0) {
		// A baseline shift is a percentage of the font size, which is what `em` is.
		style.push(`vertical-align:${props.baselinePct / 100}em`, 'font-size:0.65em')
	}

	return style.join(';')
}

/**
 * One run.
 *
 * The `contenteditable` here is the editable surface made visible: run text is
 * in surface, so the span that holds it is the span a human may type into.
 * Marking each run separately rather than making the whole frame editable is
 * what keeps an edit from merging two runs or moving text across a formatting
 * boundary — an edit the return path would have to guess at, which is exactly
 * what the surface exists to prevent.
 *
 * `data-pxh-props` carries the other four surface values — and it is the *stated*
 * ones, not the painted ones. The `style` beside it is `props.X ?? resolved.X`
 * under the frame's shrink, which cannot be read back: a placeholder title
 * painted at the layout's 44pt states no size at all, and taking 44 off the span
 * would write the layout's value into the slide — taking 30.8 off it would write
 * a number no part of the deck contains. The attribute is omitted entirely when
 * the run states none of the four, so the common case costs nothing.
 *
 * A `null` `nodeId` means the run is **chrome** — text on the slide's layout or
 * master, which is drawn and nothing more. It gets its style and neither of the
 * surface attributes, so it is not typeable and not addressable: the wordmark on
 * a template band belongs to the layout part, and honouring an edit to it would
 * change every slide in the deck.
 */
function renderRun(
	run: TextRun,
	nodeId: string | null,
	paragraph: number,
	index: number,
	fontScalePct: number | undefined
): string {
	const style = runStyle(run, fontScalePct)
	const link = run.props.hyperlink
	const styleAttr = style === '' ? '' : ` style="${escapeAttr(style)}"`
	const span =
		nodeId === null
			? `<span${styleAttr}>${escapeText(run.text)}</span>`
			: surfaceSpan(run, `${nodeId}/${paragraph}/${index}`, styleAttr)

	if (link?.url == null) return span
	return `<a href="${escapeAttr(link.url)}"${link.tooltip === undefined ? '' : ` title="${escapeAttr(link.tooltip)}"`}>${span}</a>`
}

/** The editable form of a run: its address, its invitation to type, its stated props. */
function surfaceSpan(run: TextRun, address: string, styleAttr: string): string {
	const stated = editableRunProps(run.props)
	const statedAttr =
		Object.keys(stated).length === 0 ? '' : ` data-pxh-props="${escapeAttr(JSON.stringify(stated))}"`
	return (
		`<span data-pxh-run="${escapeAttr(address)}" contenteditable="true"${statedAttr}` +
		`${styleAttr}>${escapeText(run.text)}</span>`
	)
}

/**
 * The paragraph's half of the surface: its address, and what it states.
 *
 * The `<p>` is *not* `contenteditable` and does not become so — a paragraph
 * property is set by a control, not by typing, and the runs inside it already
 * carry the invitation to type. What the address buys is the ability to name a
 * paragraph that holds no runs: a blank line is a paragraph in the source, it can
 * state an alignment like any other, and keying the surface off runs alone would
 * have left exactly those uneditable.
 *
 * A `null` `nodeId` is chrome and gets neither attribute, for the reason
 * {@link renderRun} gives — and with the same consequence, that the parser's sweep
 * for unmodeled paragraphs stays quiet instead of reporting the template's own
 * text on every slide.
 */
function paragraphSurface(paragraph: Paragraph, nodeId: string | null, index: number): string {
	if (nodeId === null) return ''
	const stated = editableParaProps(paragraph.props)
	const statedAttr =
		Object.keys(stated).length === 0 ? '' : ` data-pxh-paraprops="${escapeAttr(JSON.stringify(stated))}"`
	return ` data-pxh-para="${escapeAttr(`${nodeId}/${index}`)}"${statedAttr}`
}

/**
 * The bullet glyph, as a leading span.
 *
 * Rendered inline rather than with CSS `list-style`, because the IR states the
 * glyph, its font, its colour and its size independently of the run that follows
 * it — a real `<ul>` would style the marker from the list item and lose all four.
 *
 * A numbered bullet is drawn as `startAt + position` with the scheme's suffix
 * approximated: the seventeen `ST_TextAutonumberScheme` members differ in
 * alphabet (roman, alpha, CJK) as well as punctuation, and rendering `iv.` as
 * `4.` is a coarse picture, while renumbering the list would be a wrong one.
 * Which is why {@link Bullet.startAt} is honoured exactly and only the glyph's
 * shape is approximate.
 */
function renderBullet(bullet: Bullet, ordinal: number): string {
	if (bullet.kind === 'none') return ''

	// Every remaining arm carries the glyph's own `a:buFont`/`a:buClr`/`a:buSzPct`,
	// which style the bullet rather than the run after it.
	const style: string[] = ['margin-right:0.35em']
	if (bullet.font !== undefined) style.push(`font-family:${JSON.stringify(bullet.font)}`)
	if (bullet.color !== undefined) style.push(`color:${cssColor(bullet.color)}`)
	if (bullet.sizePct !== undefined) style.push(`font-size:${bullet.sizePct}%`)
	const attrs = ` style="${escapeAttr(style.join(';'))}"`

	if (bullet.kind === 'character') return `<span${attrs}>${escapeText(bullet.char)}</span>`
	if (bullet.kind === 'number') {
		const start = bullet.startAt ?? 1
		return `<span${attrs} data-pxh-approx="bullet:scheme">${start + ordinal}.</span>`
	}
	return `<img data-pxh-asset="${escapeAttr(bullet.asset.$asset)}"${attrs} alt=""/>`
}

/**
 * One paragraph's box.
 *
 * `reductionPct` is the frame's baked `lnSpcReduction`, and ECMA-376 §21.1.2.1.3
 * is unusually specific about it: it is *subtracted from* the spacing rather than
 * scaling it, and it "applies only to paragraphs with percentage line spacing".
 * So a paragraph spaced in points is left alone — the one such paragraph in the
 * measured corpus would otherwise tighten by a fifth for no reason the file
 * states.
 */
function paragraphStyle(props: ParagraphProperties, reductionPct: number): string {
	const style = ['margin:0']
	if (props.align !== undefined) style.push(`text-align:${props.align === 'justify' ? 'justify' : props.align}`)
	// `@lvl` is an outline depth, and the list style it indexes into is not in the
	// IR; a flat indent per level is the honest approximation of "deeper".
	if (props.level > 0) style.push(`margin-left:${px(props.level * 18)}px`)
	if (props.marginLeftPt !== undefined) style.push(`padding-left:${px(props.marginLeftPt)}px`)
	if (props.indentPt !== undefined) style.push(`text-indent:${px(props.indentPt)}px`)
	if (props.spaceBeforePt !== undefined) style.push(`margin-top:${px(props.spaceBeforePt)}px`)
	if (props.spaceAfterPt !== undefined) style.push(`margin-bottom:${px(props.spaceAfterPt)}px`)
	if (props.lineSpacing !== undefined) {
		style.push(
			props.lineSpacing.type === 'percent'
				? `line-height:${tidy(Math.max(0, props.lineSpacing.percent - reductionPct))}%`
				: `line-height:${px(props.lineSpacing.valuePt)}px`
		)
	} else if (reductionPct > 0) {
		// The paragraph inherits its spacing, so there is no stated base to subtract
		// from and CSS cannot express one. Unitless, so it recomputes against each
		// run's own size instead of freezing to the paragraph's.
		style.push(`line-height:${tidy((SINGLE_LINE_HEIGHT * (100 - reductionPct)) / 100)}`)
	}
	return style.join(';')
}

/**
 * The ordinal a numbered bullet shows: how many numbered paragraphs at the same
 * level precede it, stopping at anything that is not one. A heading between two
 * lists restarts the count, which is what PowerPoint does.
 */
function ordinalsOf(paragraphs: readonly Paragraph[]): number[] {
	const ordinals: number[] = []
	const runningByLevel = new Map<number, number>()
	for (const paragraph of paragraphs) {
		const bullet = paragraph.props.bullet
		const level = paragraph.props.level
		if (bullet?.kind !== 'number') {
			runningByLevel.delete(level)
			ordinals.push(0)
			continue
		}
		const next = runningByLevel.get(level) ?? 0
		ordinals.push(next)
		runningByLevel.set(level, next + 1)
	}
	return ordinals
}

/**
 * The text frame's HTML, for placing inside a `<foreignObject>` sized to the
 * shape's box.
 *
 * ## A baked shrink is applied; an unbaked one is not
 *
 * A `normAutofit` frame may carry the `fontScale` and `lnSpcReduction`
 * PowerPoint arrived at, and where it does they are honoured. That is not a
 * measurement — it is two numbers the file states, no different from reading a
 * font size. What this renderer still will not do is *compute* a scale for a
 * frame that bakes none, because that needs line breaking, which would make the
 * picture depend on the machine that drew it.
 *
 * The two cases are not close. Across the measured corpus 39 frames autofit by
 * shrinking; 30 bake a scale, down to 40%, and painting those at nominal size
 * drew text two and a half times too large — the single largest local error the
 * preview had. The other 7 bake nothing, and PowerPoint draws *those* at full
 * size too until the next edit, so leaving them alone is agreement rather than
 * omission.
 *
 * `nodeId` is `null` for a chrome frame — see {@link renderRun}.
 */
export function renderTextBody(text: TextBody, nodeId: string | null): string {
	const ordinals = ordinalsOf(text.paragraphs)
	const scale = text.autofitFontScalePct
	const reduction = text.autofitLineSpaceReductionPct ?? 0
	const body = text.paragraphs
		.map((paragraph, index) => {
			const bullet = paragraph.props.bullet
			const glyph = bullet === undefined ? '' : renderBullet(bullet, ordinals[index] ?? 0)
			const runs = paragraph.runs.map((run, runIndex) => renderRun(run, nodeId, index, runIndex, scale)).join('')
			// A paragraph with no runs is a blank line in the source and has to stay
			// one: an empty `<p>` collapses to nothing without something to give it
			// height.
			const content = runs === '' ? '<br/>' : runs
			return (
				`<p${paragraphSurface(paragraph, nodeId, index)}` +
				` style="${escapeAttr(paragraphStyle(paragraph.props, reduction))}">${glyph}${content}</p>`
			)
		})
		.join('')

	const frame = [
		'height:100%',
		'box-sizing:border-box',
		'display:flex',
		'flex-direction:column',
		`justify-content:${ANCHOR[text.anchor]}`,
		`padding:${px(text.insetsPt.top)}px ${px(text.insetsPt.right)}px ${px(text.insetsPt.bottom)}px ${px(text.insetsPt.left)}px`,
		`white-space:${text.wrap ? 'pre-wrap' : 'pre'}`,
		'overflow:hidden',
	]
	// `@vert` rotates the whole frame's writing direction. The two common values
	// are a quarter turn each way; the East-Asian and WordArt modes are marked
	// approximate rather than guessed at.
	if (text.vertical === 'vert') frame.push('writing-mode:vertical-rl')
	if (text.vertical === 'vert270') frame.push('writing-mode:vertical-rl', 'transform:rotate(180deg)')

	const declared: string[] = []
	if (text.vertical === 'eaVert' || text.vertical === 'wordArtVert') declared.push('text:vertical')
	// The scale is exact and says nothing; the reduction only says something when it
	// had no stated spacing to subtract from and fell back to `SINGLE_LINE_HEIGHT`.
	if (reduction > 0 && text.paragraphs.some((paragraph) => paragraph.props.lineSpacing === undefined))
		declared.push('text:linespace')
	const approx = declared.length === 0 ? '' : ` data-pxh-approx="${declared.join(' ')}"`

	return (
		`<div xmlns="http://www.w3.org/1999/xhtml" class="pxh-text"${approx} style="${escapeAttr(frame.join(';'))}">` +
		`${body}</div>`
	)
}
