import { describe, expect, it } from 'vitest'
import type { RenderIr, ShapeNode, TableNode } from '../../src/ir/render'
import { EDITABLE_PARA_PROPS, EDITABLE_RUN_PROPS, EDITABLE_SURFACE, freeze, project } from '../../src/ir/surface'
import { SAMPLE_IR } from '../fixtures/render-ir'

const clone = (ir: RenderIr): RenderIr => structuredClone(ir)

describe('the surface is one list, not two opinions', () => {
	it('names every editable run property exactly once', () => {
		const props = EDITABLE_SURFACE.filter((rule) => rule.kind === 'runProp').map((rule) =>
			rule.path.slice(rule.path.lastIndexOf('.') + 1)
		)
		expect(props).toEqual([...EDITABLE_RUN_PROPS])
	})

	it('names every editable paragraph property exactly once', () => {
		const props = EDITABLE_SURFACE.filter((rule) => rule.kind === 'paraProp').map((rule) =>
			rule.path.slice(rule.path.lastIndexOf('.') + 1)
		)
		expect(props).toEqual([...EDITABLE_PARA_PROPS])
	})

	it('keeps `level` out, because nothing in the paint model can say what it indexes into', () => {
		// The paragraph property most likely to be added by analogy and the one that
		// must not be: `@lvl` is an index into a list style the paint model does not
		// carry, so a control for it would be offering to change a number whose meaning
		// lives somewhere this package never reads.
		expect(EDITABLE_SURFACE.some((rule) => rule.path.endsWith('.level'))).toBe(false)
		expect(EDITABLE_PARA_PROPS).not.toContain('level')
	})

	it('permits run text and node deletion, and nothing structural', () => {
		expect(EDITABLE_SURFACE.filter((rule) => rule.kind === 'text')).toHaveLength(1)
		expect(EDITABLE_SURFACE.filter((rule) => rule.kind === 'delete')).toHaveLength(1)
		// Moving a box, changing geometry or reordering slides must stay out: each
		// needs interpretation on the way back, and the return path never guesses.
		expect(EDITABLE_SURFACE.some((rule) => /placement|geometry|slides/.test(rule.path))).toBe(false)
	})
})

describe('project', () => {
	const projection = project(SAMPLE_IR)

	it('reaches runs wherever they live — shapes, table cells, group children', () => {
		const owners = new Set(
			projection.slides.flatMap((slide) => slide.nodes.flatMap((node) => node.runs.map((run) => run.node)))
		)
		expect(owners).toContain('s1.sp2') // the title shape
		expect(owners).toContain('s1.sp8') // a shape nested inside a group
		expect(owners).toContain('s2.sp2.r0c0') // a table cell
	})

	it('addresses a run by node id, not by position', () => {
		// Node deletion is in surface, so anything positional would misalign the
		// moment a user removes a shape — the case the surface explicitly allows.
		const title = projection.slides[0]?.nodes.find((node) => node.id === 's1.sp2')
		expect(title?.runs.map((run) => run.text)).toEqual(['Quarterly ', 'review'])
		expect(title?.runs.map((run) => run.paragraph)).toEqual([0, 0])
		expect(title?.runs.map((run) => run.run)).toEqual([0, 1])
	})

	it('carries only the sanctioned properties', () => {
		// The title's second run is underlined and states a character spacing. The
		// first is in surface — `RunProperties.underline` models the three values the
		// write API expresses and nothing else — and `spacingPt` is not, so the
		// editable view must show one and not the other.
		const styled = projection.slides[0]?.nodes.find((node) => node.id === 's1.sp2')?.runs[1]
		expect(styled?.props).toStrictEqual({
			sizePt: 40,
			italic: true,
			underline: 'single',
			// A theme reference reaches the editable view intact rather than
			// flattened to the hex it happened to resolve to, so an edit that does
			// not touch the colour cannot silently break the theme binding.
			color: {
				kind: 'scheme',
				slot: 'accent1',
				transforms: [{ name: 'lumMod', value: '75000' }],
				effectiveHex: '2E5A8A',
			},
		})
	})

	it('omits a property the run does not state, rather than writing undefined', () => {
		// An absent character property means *inherited*. Spelling it `undefined`
		// would both break the island's "absence has one spelling" rule and turn
		// inheritance into an explicit value on the way back.
		const plain = projection.slides[0]?.nodes.find((node) => node.id === 's1.sp3')?.runs[0]
		expect(plain?.props).toStrictEqual({})
		// The JSON round trip is the assertion, not a way to copy: `structuredClone`
		// here would prove that the model survives structured cloning, which is not
		// the claim this test makes.
		// oxlint-disable-next-line unicorn/prefer-structured-clone
		expect(JSON.parse(JSON.stringify(projection))).toStrictEqual(projection)
	})

	it('carries a paragraph’s alignment, bullet and margins, and only what the paragraph states', () => {
		const bullets = projection.slides[0]?.nodes.find((node) => node.id === 's1.sp3')
		// Three paragraphs: one stating an alignment, a bullet and both margins, one
		// stating a bullet and no alignment, and a blank line stating an alignment and
		// no bullet. `level` and `spaceBeforePt` sit beside them on the same object and
		// are not in surface, so a projection that copied the props wholesale would fail
		// on the first of these.
		expect(bullets?.paragraphs.map((paragraph) => paragraph.props)).toStrictEqual([
			{
				align: 'left',
				bullet: { kind: 'character', char: '•', font: 'Arial' },
				marginLeftPt: 18,
				indentPt: -18,
			},
			{ bullet: { kind: 'number', scheme: 'arabicPeriod', startAt: 1 } },
			{ align: 'right' },
		])
	})

	it('carries the bullets the write API cannot author back, because carrying is not offering', () => {
		// The two above are a glyph with its own font and a numbering scheme, and
		// neither is settable — `parse/edits.ts` refuses both rather than approximating.
		// They are projected anyway, and must be: the projection is what the renderer
		// writes onto the `<p>` and what comes back, so a bullet left out of it would
		// read as *cleared* on the return path and be written as `'inherit'`, silently
		// stripping a glyph from every deck that has one.
		const bullets = projection.slides[0]?.nodes.find((node) => node.id === 's1.sp3')
		expect(bullets?.paragraphs[1]?.props.bullet).toStrictEqual({
			kind: 'number',
			scheme: 'arabicPeriod',
			startAt: 1,
		})
	})

	it('projects a paragraph that holds no runs', () => {
		// A blank line is a paragraph in the source and can state an alignment like
		// any other. Keying the surface off runs would have made exactly the empty
		// paragraphs uneditable — and silently, since they have nothing to show for it.
		const bullets = projection.slides[0]?.nodes.find((node) => node.id === 's1.sp3')
		expect(bullets?.paragraphs).toHaveLength(3)
		expect(bullets?.runs.filter((run) => run.paragraph === 2)).toHaveLength(0)
		expect(bullets?.paragraphs[2]).toStrictEqual({ node: 's1.sp3', paragraph: 2, props: { align: 'right' } })
	})

	it('flattens groups but keeps opaque and picture nodes visible with no runs', () => {
		const kinds = projection.slides.flatMap((slide) => slide.nodes.map((node) => node.kind))
		expect(kinds).not.toContain('group')
		expect(kinds).toContain('opaque')
		expect(kinds).toContain('picture')
	})
})

describe('freeze is the complement of project', () => {
	it('is unchanged by an in-surface edit', () => {
		// This is the load-bearing property. The parser hashes the frozen model to
		// decide whether anything *outside* the surface moved, so a legitimate
		// text or bold edit has to leave it bit-identical.
		const edited = clone(SAMPLE_IR)
		const title = edited.slides[0]?.nodes[0] as ShapeNode
		const run = title.text?.paragraphs[0]?.runs[0]
		if (!run) throw new Error('fixture lost its title run')
		run.text = 'Annual review'
		run.props.bold = true
		delete run.props.color

		expect(freeze(edited)).toStrictEqual(freeze(SAMPLE_IR))
	})

	it('changes when an out-of-surface edit moves a box', () => {
		const drifted = clone(SAMPLE_IR)
		const title = drifted.slides[0]?.nodes[0]
		if (!title?.placement) throw new Error('fixture lost its title node')
		title.placement.box.x += 914_400

		expect(freeze(drifted)).not.toStrictEqual(freeze(SAMPLE_IR))
	})

	it('changes when an out-of-surface character property moves', () => {
		// `spacingPt` is a run property like `bold`, and the only thing separating
		// them is this list. If `freeze` stripped by shape rather than by name,
		// this would pass silently and emit would accept an edit it cannot apply.
		const drifted = clone(SAMPLE_IR)
		const title = drifted.slides[0]?.nodes[0] as ShapeNode
		const run = title.text?.paragraphs[0]?.runs[1]
		if (!run) throw new Error('fixture lost its styled run')
		run.props.spacingPt = 3

		expect(freeze(drifted)).not.toStrictEqual(freeze(SAMPLE_IR))
	})

	it('is unchanged when an underline moves, now that underline is in surface', () => {
		// The mirror of the case above, and the reason both are here: the two
		// properties sit side by side on the same run and differ only in being named
		// in `EDITABLE_RUN_PROPS`. A `freeze` that stripped by shape — "every string
		// on a run" — would pass the test above and fail this one.
		const edited = clone(SAMPLE_IR)
		const title = edited.slides[0]?.nodes[0] as ShapeNode
		const run = title.text?.paragraphs[0]?.runs[1]
		if (!run) throw new Error('fixture lost its styled run')
		run.props.underline = 'double'
		run.props.strike = 'single'

		expect(freeze(edited)).toStrictEqual(freeze(SAMPLE_IR))
	})

	it('is unchanged when a paragraph’s alignment, bullet or margins move, and changes when its level does', () => {
		// The paragraph tier's version of the pair above, and the same trap: `align`,
		// `bullet`, the two margins, `level` and `spaceBeforePt` are all neighbours on
		// one object, and only the first four are in surface. A `freeze` that stripped
		// `props` wholesale would pass the first half here and let an out-of-surface
		// outline level change go unnoticed.
		const edited = clone(SAMPLE_IR)
		const bullets = edited.slides[0]?.nodes.find((node) => node.id === 's1.sp3') as ShapeNode
		const paragraph = bullets.text?.paragraphs[0]
		if (!paragraph) throw new Error('fixture lost its bulleted paragraph')
		paragraph.props.align = 'justify'
		paragraph.props.bullet = { kind: 'none' }
		paragraph.props.marginLeftPt = 72
		delete paragraph.props.indentPt
		expect(freeze(edited)).toStrictEqual(freeze(SAMPLE_IR))

		const drifted = clone(SAMPLE_IR)
		const sameNode = drifted.slides[0]?.nodes.find((node) => node.id === 's1.sp3') as ShapeNode
		const other = sameNode.text?.paragraphs[0]
		if (!other) throw new Error('fixture lost its bulleted paragraph')
		other.props.level = 3
		expect(freeze(drifted)).not.toStrictEqual(freeze(SAMPLE_IR))

		// And the neighbour that is a measurement like the margins are, so the
		// separation cannot be "strip every number on a paragraph".
		const spaced = clone(SAMPLE_IR)
		const spacedNode = spaced.slides[0]?.nodes.find((node) => node.id === 's1.sp3') as ShapeNode
		const spacedParagraph = spacedNode.text?.paragraphs[0]
		if (!spacedParagraph) throw new Error('fixture lost its bulleted paragraph')
		spacedParagraph.props.spaceBeforePt = 24
		expect(freeze(spaced)).not.toStrictEqual(freeze(SAMPLE_IR))
	})

	it('strips text inside table cells and groups too', () => {
		const frozen = freeze(SAMPLE_IR)
		const table = frozen.slides[1]?.nodes[0] as TableNode
		// `resolved` survives the freeze on purpose: it is what import computed for
		// painting, nothing on the editable surface writes it, and a renderer that
		// mangled it would otherwise change the page with no drift to show for it.
		expect(table.rows[0]?.cells[0]?.text?.paragraphs[0]?.runs[0]).toStrictEqual({
			text: '',
			props: {},
			resolved: { bold: true, sizePt: 18 },
		})

		const group = frozen.slides[0]?.nodes[5]
		if (group?.kind !== 'group') throw new Error('fixture lost its group')
		const nested = group.children[0] as ShapeNode
		expect(nested.text?.paragraphs[0]?.runs[0]?.text).toBe('')
	})

	it('leaves a deleted node’s siblings untouched', () => {
		// Deletion is in surface, so the frozen node lists are compared as a
		// subsequence rather than for equality. What must hold is that removing
		// one node changes nothing about the others.
		const reduced = clone(SAMPLE_IR)
		const slide = reduced.slides[0]
		if (!slide) throw new Error('fixture lost its first slide')
		slide.nodes.splice(2, 1)

		const before = freeze(SAMPLE_IR).slides[0]?.nodes ?? []
		const after = freeze(reduced).slides[0]?.nodes ?? []
		expect(after).toStrictEqual(before.filter((node) => node.id !== 's1.sp4'))
	})
})
