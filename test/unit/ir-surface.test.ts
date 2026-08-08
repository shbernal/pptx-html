import { describe, expect, it } from 'vitest'
import type { RenderIr, ShapeNode, TableNode } from '../../src/ir/render'
import { EDITABLE_RUN_PROPS, EDITABLE_SURFACE, freeze, project } from '../../src/ir/surface'
import { SAMPLE_IR } from '../fixtures/render-ir'

const clone = (ir: RenderIr): RenderIr => JSON.parse(JSON.stringify(ir)) as RenderIr

describe('the surface is one list, not two opinions', () => {
	it('names every editable run property exactly once', () => {
		const props = EDITABLE_SURFACE.filter((rule) => rule.kind === 'runProp').map((rule) =>
			rule.path.slice(rule.path.lastIndexOf('.') + 1)
		)
		expect(props).toEqual([...EDITABLE_RUN_PROPS])
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
		// The title's second run is underlined. `underline` is deliberately out of
		// surface — `ST_TextUnderlineType` has seventeen members and the write API
		// expresses three — so it must not appear in the editable view.
		const styled = projection.slides[0]?.nodes.find((node) => node.id === 's1.sp2')?.runs[1]
		expect(styled?.props).toStrictEqual({
			sizePt: 40,
			italic: true,
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
		expect(JSON.parse(JSON.stringify(projection))).toStrictEqual(projection)
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
		// This is the load-bearing property. Part 06 hashes the frozen model to
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
		if (!title) throw new Error('fixture lost its title node')
		title.placement.box.x += 914_400

		expect(freeze(drifted)).not.toStrictEqual(freeze(SAMPLE_IR))
	})

	it('changes when an out-of-surface character property moves', () => {
		// `underline` is a run property like `bold`, and the only thing separating
		// them is this list. If `freeze` stripped by shape rather than by name,
		// this would pass silently and part 06 would accept an edit it cannot emit.
		const drifted = clone(SAMPLE_IR)
		const title = drifted.slides[0]?.nodes[0] as ShapeNode
		const run = title.text?.paragraphs[0]?.runs[1]
		if (!run) throw new Error('fixture lost its styled run')
		run.props.underline = 'double'

		expect(freeze(drifted)).not.toStrictEqual(freeze(SAMPLE_IR))
	})

	it('strips text inside table cells and groups too', () => {
		const frozen = freeze(SAMPLE_IR)
		const table = frozen.slides[1]?.nodes[0] as TableNode
		expect(table.rows[0]?.cells[0]?.text?.paragraphs[0]?.runs[0]).toStrictEqual({ text: '', props: {} })

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
