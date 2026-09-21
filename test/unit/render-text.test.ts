/**
 * The baked-shrink arms, on hand-built frames.
 *
 * The corpus reaches the two that matter most — a scale applied to a run, and a
 * reduction against a stated percentage — but not the exemption ECMA-376
 * §21.1.2.1.3 attaches to the reduction: *"this attribute applies only to
 * paragraphs with percentage line spacing"*. The write API can author a frame
 * that hits it, so this is not a corpus gap to close later; it is a rule read off
 * the spec rather than observed, and a rule nothing exercises is a rule nobody
 * will notice breaking.
 */

import { describe, expect, it } from 'vitest'
import type { ParagraphProperties, TextBody } from '../../src/ir/render'
import { renderTextBody } from '../../src/render/text'

function frame(spacing: ParagraphProperties['lineSpacing'], autofit: Partial<TextBody>): TextBody {
	return {
		paragraphs: [
			{
				props: { level: 0, ...(spacing === undefined ? {} : { lineSpacing: spacing }) },
				runs: [{ text: 'x', props: { sizePt: 20 }, resolved: { sizePt: 20 } }],
			},
		],
		autofit: 'shrink',
		anchor: 'top',
		wrap: true,
		insetsPt: { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 },
		...autofit,
	}
}

describe('a baked shrink', () => {
	it('scales the run size by the stated percentage', () => {
		expect(renderTextBody(frame(undefined, { autofitFontScalePct: 62.5 }), 'n')).toContain('font-size:12.5px')
	})

	it('leaves a frame that bakes no scale at its nominal size', () => {
		// `shrink` with no numbers is a frame PowerPoint has not measured; it draws
		// those at full size too, until the next edit.
		const html = renderTextBody(frame(undefined, {}), 'n')
		expect(html).toContain('font-size:20px')
		expect(html).not.toContain('data-pxh-approx')
	})

	it('subtracts the reduction from a percentage spacing rather than scaling it', () => {
		// 150% less 20 points of reduction is 130%, not 120% — the spec says
		// subtracted, and the two only agree at 100%.
		//
		// Written unitless, as `1.3`. PowerPoint's percentage is a multiple of each line's
		// own font size, and a CSS percentage is not: it resolves once against the element
		// it is written on, which is the paragraph, while the size lives on the runs. The
		// arithmetic is unchanged; only the spelling is.
		const html = renderTextBody(frame({ type: 'percent', percent: 150 }, { autofitLineSpaceReductionPct: 20 }), 'n')
		expect(html).toContain('line-height:1.3')
		// Nothing was guessed at: the paragraph stated its own base.
		expect(html).not.toContain('data-pxh-approx')
	})

	it('leaves a spacing stated in points alone', () => {
		// The exemption. A 24pt line stays 24pt under a 20% reduction, because the
		// attribute reaches percentage spacing only.
		const html = renderTextBody(frame({ type: 'points', valuePt: 24 }, { autofitLineSpaceReductionPct: 20 }), 'n')
		expect(html).toContain('line-height:24px')
		expect(html).not.toContain('data-pxh-approx')
	})

	it('falls back to a stand-in single spacing when the paragraph states none, and declares it', () => {
		// The one approximate arm: CSS cannot say "`normal`, less a fifth", so the
		// base is a documented constant and the frame is marked.
		const html = renderTextBody(frame(undefined, { autofitLineSpaceReductionPct: 20 }), 'n')
		expect(html).toContain('line-height:0.96')
		expect(html).toContain('data-pxh-approx="text:linespace"')
	})

	it('never lets a reduction drive the line height negative', () => {
		const html = renderTextBody(frame({ type: 'percent', percent: 80 }, { autofitLineSpaceReductionPct: 100 }), 'n')
		// Matched to the end of the declaration: a bare `line-height:0` is a prefix of the
		// `0.96` stand-in two cases up, so it would pass on the wrong value.
		expect(html).toMatch(/line-height:0[;"]/)
	})
})
