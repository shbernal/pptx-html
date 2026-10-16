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
	it('leaves a frame that bakes no scale at its nominal size', () => {
		// `shrink` with no numbers is a frame PowerPoint has not measured; it draws
		// those at full size too, until the next edit.
		const html = renderTextBody(frame(undefined, {}), 'n')
		expect(html).toContain('font-size:20px')
		expect(html).not.toContain('data-pxh-approx')
	})

	it('leaves a spacing stated in points alone', () => {
		// The exemption. A 24pt line stays 24pt under a 20% reduction, because the
		// attribute reaches percentage spacing only.
		const html = renderTextBody(frame({ type: 'points', valuePt: 24 }, { autofitLineSpaceReductionPct: 20 }), 'n')
		expect(html).toContain('line-height:24px')
		expect(html).not.toContain('data-pxh-approx')
	})

	it('never lets a reduction drive the line height negative', () => {
		const html = renderTextBody(frame({ type: 'percent', percent: 80 }, { autofitLineSpaceReductionPct: 100 }), 'n')
		// Matched to the end of the declaration: a bare `line-height:0` is also a prefix
		// of a fallback like `0.96`, so it would pass on the wrong value.
		expect(html).toMatch(/line-height:0[;"]/)
	})
})
