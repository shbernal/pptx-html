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
 * The slide's `viewBox` is in EMU, so one SVG user unit is one EMU, and inside a
 * `foreignObject` one CSS pixel is one user unit. Every length below is
 * therefore EMU written as `px` — a 44pt run is `font-size: 558800px`, which
 * looks alarming and is exactly right.
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
import { editableRunProps } from '../ir/surface'
import { cssColor, EMU_PER_POINT, escapeAttr } from './paint'

/** Text-node escaping. `&` first, or it would double-escape the entities below. */
export function escapeText(value: string): string {
	return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
}

const ANCHOR: Record<TextBody['anchor'], string> = {
	top: 'flex-start',
	middle: 'center',
	bottom: 'flex-end',
}

/** Points → EMU, rounded to the integer the canvas is stated in. */
function emu(points: number): number {
	return Math.round(points * EMU_PER_POINT)
}

function runStyle(run: TextRun): string {
	const props = run.props
	const resolved = run.resolved
	const style: string[] = []

	const face = props.fontFace ?? resolved.fontFace
	if (face !== undefined) style.push(`font-family:${JSON.stringify(face)},sans-serif`)

	const size = props.sizePt ?? resolved.sizePt
	if (size !== undefined) style.push(`font-size:${emu(size)}px`)

	const bold = props.bold ?? resolved.bold
	if (bold !== undefined) style.push(`font-weight:${bold ? 700 : 400}`)

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

	if (props.spacingPt !== undefined) style.push(`letter-spacing:${emu(props.spacingPt)}px`)
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
 * `data-d2p-props` carries the other four surface values — and it is the *stated*
 * ones, not the painted ones. The `style` beside it is `props.X ?? resolved.X`,
 * which cannot be read back: a placeholder title painted at the layout's 44pt
 * states no size at all, and taking 44 off the span would write the layout's
 * value into the slide. The attribute is omitted entirely when the run states
 * none of the four, so the common case costs nothing.
 */
function renderRun(run: TextRun, nodeId: string, paragraph: number, index: number): string {
	const style = runStyle(run)
	const address = `${nodeId}/${paragraph}/${index}`
	const link = run.props.hyperlink
	const stated = editableRunProps(run.props)
	const statedAttr =
		Object.keys(stated).length === 0 ? '' : ` data-d2p-props="${escapeAttr(JSON.stringify(stated))}"`
	const span =
		`<span data-d2p-run="${escapeAttr(address)}" contenteditable="true"${statedAttr}` +
		`${style === '' ? '' : ` style="${escapeAttr(style)}"`}>${escapeText(run.text)}</span>`

	if (link?.url == null) return span
	return `<a href="${escapeAttr(link.url)}"${link.tooltip === undefined ? '' : ` title="${escapeAttr(link.tooltip)}"`}>${span}</a>`
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
		return `<span${attrs} data-d2p-approx="bullet:scheme">${start + ordinal}.</span>`
	}
	return `<img data-d2p-asset="${escapeAttr(bullet.asset.$asset)}"${attrs} alt=""/>`
}

function paragraphStyle(props: ParagraphProperties): string {
	const style = ['margin:0']
	if (props.align !== undefined) style.push(`text-align:${props.align === 'justify' ? 'justify' : props.align}`)
	// `@lvl` is an outline depth, and the list style it indexes into is not in the
	// IR; a flat indent per level is the honest approximation of "deeper".
	if (props.level > 0) style.push(`margin-left:${props.level * emu(18)}px`)
	if (props.marginLeftPt !== undefined) style.push(`padding-left:${emu(props.marginLeftPt)}px`)
	if (props.indentPt !== undefined) style.push(`text-indent:${emu(props.indentPt)}px`)
	if (props.spaceBeforePt !== undefined) style.push(`margin-top:${emu(props.spaceBeforePt)}px`)
	if (props.spaceAfterPt !== undefined) style.push(`margin-bottom:${emu(props.spaceAfterPt)}px`)
	if (props.lineSpacing !== undefined) {
		style.push(
			props.lineSpacing.type === 'percent'
				? `line-height:${props.lineSpacing.percent}%`
				: `line-height:${emu(props.lineSpacing.valuePt)}px`
		)
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
 * Autofit is read but not applied. `shrink` (`normAutofit`) states a font scale
 * PowerPoint computed by measuring, and re-deriving it here would make the
 * rendered size depend on the machine that rendered it; the box is the box, and
 * text that overflows it overflows visibly rather than being silently rescaled.
 */
export function renderTextBody(text: TextBody, nodeId: string): string {
	const ordinals = ordinalsOf(text.paragraphs)
	const body = text.paragraphs
		.map((paragraph, index) => {
			const bullet = paragraph.props.bullet
			const glyph = bullet === undefined ? '' : renderBullet(bullet, ordinals[index] ?? 0)
			const runs = paragraph.runs.map((run, runIndex) => renderRun(run, nodeId, index, runIndex)).join('')
			// A paragraph with no runs is a blank line in the source and has to stay
			// one: an empty `<p>` collapses to nothing without something to give it
			// height.
			const content = runs === '' ? '<br/>' : runs
			return `<p style="${escapeAttr(paragraphStyle(paragraph.props))}">${glyph}${content}</p>`
		})
		.join('')

	const frame = [
		'height:100%',
		'box-sizing:border-box',
		'display:flex',
		'flex-direction:column',
		`justify-content:${ANCHOR[text.anchor]}`,
		`padding:${emu(text.insetsPt.top)}px ${emu(text.insetsPt.right)}px ${emu(text.insetsPt.bottom)}px ${emu(text.insetsPt.left)}px`,
		`white-space:${text.wrap ? 'pre-wrap' : 'pre'}`,
		'overflow:hidden',
	]
	// `@vert` rotates the whole frame's writing direction. The two common values
	// are a quarter turn each way; the East-Asian and WordArt modes are marked
	// approximate rather than guessed at.
	if (text.vertical === 'vert') frame.push('writing-mode:vertical-rl')
	if (text.vertical === 'vert270') frame.push('writing-mode:vertical-rl', 'transform:rotate(180deg)')

	const approx =
		text.vertical === 'eaVert' || text.vertical === 'wordArtVert' ? ' data-d2p-approx="text:vertical"' : ''

	return (
		`<div xmlns="http://www.w3.org/1999/xhtml" class="d2p-text"${approx} style="${escapeAttr(frame.join(';'))}">` +
		`${body}</div>`
	)
}
