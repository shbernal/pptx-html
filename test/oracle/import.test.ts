/**
 * What the importer must get right, checked against real decks rather than
 * hand-written input — because the mapping's failure mode is not "throws", it is
 * "produces something plausible that the source did not say".
 *
 * Lives under `test/oracle/` rather than `test/unit/` for one reason: every case
 * here needs a corpus deck built by the writer, which is the same generation the
 * oracle project owns. No DOM is involved.
 */

import { describe, expect, it } from 'vitest'
import { importDeck } from '../../src/import/deck'
import type { RenderIr, RenderNode, ShapeNode } from '../../src/ir/render'
import { CORPUS, corpusBytes } from '../corpus/decks'

const cache = new Map<string, RenderIr>()

async function importCorpus(name: string): Promise<RenderIr> {
	const hit = cache.get(name)
	if (hit) return hit
	const entry = CORPUS.find((candidate) => candidate.name === name)
	if (!entry) throw new Error(`the corpus has no deck named ${name}`)
	const { render } = await importDeck(await corpusBytes(entry))
	cache.set(name, render)
	return render
}

function flatten(nodes: readonly RenderNode[]): RenderNode[] {
	return nodes.flatMap((node) => (node.kind === 'group' ? [node, ...flatten(node.children)] : [node]))
}

async function nodeNamed(deck: string, name: string): Promise<RenderNode> {
	const ir = await importCorpus(deck)
	const node = ir.slides.flatMap((slide) => flatten(slide.nodes)).find((candidate) => candidate.name === name)
	if (!node) throw new Error(`${deck} has no node named ${name}`)
	return node
}

describe('the imported model survives the JSON island', () => {
	for (const entry of CORPUS) {
		it(`${entry.name} imports to a JSON-native model`, async () => {
			// The same rule `test/unit/ir-render.test.ts` holds the hand-written
			// fixture to, applied to what the importer really emits: no `undefined`,
			// no `Uint8Array`, no exotic prototypes, no non-finite numbers. A mapper
			// that returned a `Buffer` or wrote an explicit `undefined` would pass
			// every structural assertion below and fail only on the way back.
			const ir = await importCorpus(entry.name)
			const offenders: string[] = []
			walk(ir, '$', offenders)
			expect(offenders).toEqual([])
			expect(JSON.parse(JSON.stringify(ir))).toStrictEqual(ir)
		})
	}
})

function walk(value: unknown, path: string, offenders: string[]): void {
	if (value === null || typeof value === 'boolean' || typeof value === 'string') return
	if (typeof value === 'number') {
		if (!Number.isFinite(value)) offenders.push(`${path}: ${value}`)
		return
	}
	if (Array.isArray(value)) {
		for (const [index, entry] of value.entries()) walk(entry, `${path}[${index}]`, offenders)
		return
	}
	if (typeof value !== 'object') {
		offenders.push(`${path}: ${typeof value}`)
		return
	}
	if (Object.getPrototypeOf(value) !== Object.prototype) {
		offenders.push(`${path}: ${value.constructor?.name ?? 'exotic object'}`)
		return
	}
	for (const [key, entry] of Object.entries(value)) walk(entry, `${path}.${key}`, offenders)
}

describe('identity', () => {
	it('is stable across two imports of the same bytes', async () => {
		// The property the whole differ depends on: ids are derived from the source,
		// so a second import agrees with the first. A generated id would make every
		// comparison the return path makes vacuous.
		const bytes = await corpusBytes(CORPUS[0])
		const first = await importDeck(bytes)
		const second = await importDeck(bytes)
		expect(second.render).toStrictEqual(first.render)
	})

	it('is unique across a deck, including nested group children and table cells', async () => {
		for (const entry of CORPUS) {
			const ir = await importCorpus(entry.name)
			const ids: string[] = []
			for (const slide of ir.slides) {
				for (const node of flatten(slide.nodes)) {
					ids.push(node.id)
					if (node.kind === 'table') for (const row of node.rows) for (const cell of row.cells) ids.push(cell.id)
				}
			}
			expect(new Set(ids).size, `${entry.name} has duplicate node ids`).toBe(ids.length)
		}
	})
})

describe('geometry is slide-absolute at every depth', () => {
	it('places a doubly nested child in slide space, not its group’s', async () => {
		// The trap: coordinates inside `p:grpSp` are stated in the group's child
		// space. Reading them as slide coordinates displaces every grouped shape,
		// and the error compounds with depth — which is why the corpus has a group
		// inside a group and why this checks the *inner* child.
		const inner = await nodeNamed('nested-group', 'inner-plate')
		const outer = await nodeNamed('nested-group', 'outer')
		expect(inner.placement).not.toBeNull()
		expect(outer.placement).not.toBeNull()
		if (!inner.placement || !outer.placement) return

		// 2" from the slide origin, as authored — not 2" from the enclosing group.
		expect(inner.placement.box.x).toBe(1_828_800)
		expect(inner.placement.box.y).toBe(731_520)
		expect(inner.placement.box.x).toBeGreaterThan(outer.placement.box.x)
	})

	it('records rotation and flips rather than folding them into the box', async () => {
		const rotated = await nodeNamed('rotated-flipped', 'rotated')
		const flipped = await nodeNamed('rotated-flipped', 'flipped')
		expect(rotated.placement?.rotation).toBe(30)
		expect(rotated.placement?.flipH).toBe(false)
		expect(flipped.placement?.flipH).toBe(true)
		expect(flipped.placement?.flipV).toBe(true)
		// The box is PowerPoint's unrotated placement box; a renderer applies the
		// rotation. Baking it in would make the box unwritable back to a deck.
		expect(flipped.placement?.box.w).toBe(1_828_800)
	})
})

describe('a colour keeps its theme reference', () => {
	it('imports a scheme fill as a token plus the pixel it resolves to', async () => {
		// Flattening to the hex passes a pixel diff and loses the reference, so a
		// re-themed deck stops re-colouring. Both halves have to be there.
		const themed = (await nodeNamed('theme-color', 'themed')) as ShapeNode
		expect(themed.fill).toStrictEqual({
			kind: 'solid',
			color: { kind: 'scheme', slot: 'accent1', transforms: [], effectiveHex: '250F6B' },
		})
	})

	it('imports an explicit colour as srgb, with no invented token', async () => {
		const card = (await nodeNamed('autoshape', 'card')) as ShapeNode
		expect(card.fill).toStrictEqual({ kind: 'solid', color: { kind: 'srgb', hex: 'DDE3F0' } })
	})
})

describe('absent means inherited', () => {
	it('leaves a placeholder run’s properties unstated and resolves them separately', async () => {
		// The flattening trap in miniature. This run states no size, face or colour —
		// it takes all three from the layout's placeholder. Writing 44pt into `props`
		// would render identically and bake the layout's size into the slide the
		// moment emit read it back.
		const title = (await nodeNamed('layout-placeholder', 'title')) as ShapeNode
		const run = title.text?.paragraphs[0]?.runs[0]
		expect(run?.props).toStrictEqual({})
		expect(run?.resolved.sizePt).toBe(44)
		expect(run?.resolved.fontFace).toBe('Calibri Light')
	})

	it('keeps a stated property in props and mirrors it into resolved', async () => {
		const copy = (await nodeNamed('text-box', 'copy')) as ShapeNode
		const run = copy.text?.paragraphs[0]?.runs[0]
		expect(run?.props.bold).toBe(true)
		expect(run?.props.sizePt).toBe(24)
		expect(run?.resolved.bold).toBe(true)
	})

	it('distinguishes an explicit no-line from an unstated one', async () => {
		const freeform = (await nodeNamed('custgeom', 'freeform')) as ShapeNode
		const card = (await nodeNamed('autoshape', 'card')) as ShapeNode
		const copy = (await nodeNamed('text-box', 'copy')) as ShapeNode
		expect(freeform.stroke.kind).toBe('none')
		expect(card.stroke).toStrictEqual({
			kind: 'line',
			widthPt: 2,
			color: { kind: 'srgb', hex: '250F6B' },
			dash: 'solid',
		})
		expect(copy.stroke.kind).toBe('inherit')
	})

	it('distinguishes an explicit no-fill from an unstated one', async () => {
		// Before `Shape.fillNoFill` (ts-pptx 3.0.0, closing upstream #1) this reported
		// `inherit`, and so did every other text box in the corpus — the model said
		// "takes its fill from the style reference" about a box whose XML says
		// `a:noFill`. That is the flattening the two arms exist to prevent, and it was
		// the corpus-wide default until the accessor landed.
		const copy = (await nodeNamed('text-box', 'copy')) as ShapeNode
		const card = (await nodeNamed('autoshape', 'card')) as ShapeNode
		expect(copy.fill).toStrictEqual({ kind: 'none' })
		expect(card.fill).toStrictEqual({ kind: 'solid', color: { kind: 'srgb', hex: 'DDE3F0' } })

		// The third arm has no corpus case on purpose rather than by omission: a fill
		// of `inherit` needs a shape with a `p:style/a:fillRef` and no fill child, and
		// this writer emits a fill or an `a:noFill` on everything it authors. Since
		// ts-pptx 3.1.0 that is true of *every* `addShape` call — `fill: { type: 'none' }`
		// now emits `<a:noFill/>` as it always said it did (upstream #9), and omitting
		// `fill` emits it too, so the write API has no spelling for `inherit` at all
		// (https://github.com/shbernal/ts-pptx/issues/10). The arm is reachable from
		// the DOM lane and from PowerPoint-authored decks, which is the deferred second
		// corpus tier.
		const fills = (await importCorpus('autoshape')).slides
			.flatMap((slide) => flatten(slide.nodes))
			.filter((node) => node.kind === 'shape')
			.map((node) => node.fill.kind)
		expect(fills).not.toContain('inherit')
	})

	it('keeps a stated line cap and leaves an unstated one absent', async () => {
		// `@cap` is modeled because `ShapeLineProps.cap` can carry it back; `@algn` is
		// a note instead, because nothing can. Neither is a default this file invents:
		// an unstated cap is an absent key, not `'flat'`.
		const rule = (await nodeNamed('line-cap', 'rule')) as ShapeNode
		const card = (await nodeNamed('autoshape', 'card')) as ShapeNode
		expect(rule.stroke).toStrictEqual({
			kind: 'line',
			widthPt: 6,
			color: { kind: 'srgb', hex: '250F6B' },
			dash: 'dash',
			cap: 'rnd',
		})
		expect(card.stroke).not.toHaveProperty('cap')
	})
})

describe('bullets', () => {
	it('keeps a numbered list’s startAt and the bullet’s own font, size and colour', async () => {
		// All four were unreachable before `bulletDetail` replaced the tagged string
		// (upstream #3). `startAt` is the one that changes what the slide *says*.
		const steps = (await nodeNamed('bullet', 'steps')) as ShapeNode
		expect(steps.text?.paragraphs[0]?.props.bullet).toStrictEqual({
			kind: 'number',
			scheme: 'arabicPeriod',
			startAt: 5,
			font: 'Wingdings',
			color: { kind: 'srgb', hex: 'C00000' },
			sizePct: 80,
		})
	})

	it('resolves a picture bullet to a hash-addressed asset, never inline bytes', async () => {
		// The fourth kind, which the tagged string dropped to `null` outright. It is
		// addressed the way every other picture in this model is — an `AssetRef`, so
		// the JSON island stays free of `Uint8Array` and the image is one entry rather
		// than one per paragraph that uses it.
		const starred = (await nodeNamed('bullet', 'starred')) as ShapeNode
		const bullet = starred.text?.paragraphs[0]?.props.bullet
		expect(bullet?.kind).toBe('picture')
		if (bullet?.kind !== 'picture') return
		const ir = await importCorpus('bullet')
		expect(ir.assets.map((asset) => asset.name)).toContain(bullet.asset.$asset)
	})

	it('leaves an inherited bullet absent and an explicit a:buNone stated', async () => {
		// The distinction the `null` arm exists for: a paragraph inheriting the list
		// style's bullet and one suppressing its own render differently, so `absent`
		// and `{ kind: 'none' }` may never collapse together.
		const copy = (await nodeNamed('text-box', 'copy')) as ShapeNode
		const bullets = copy.text?.paragraphs.map((paragraph) => paragraph.props.bullet)
		expect(bullets?.every((bullet) => bullet === undefined || bullet.kind === 'none')).toBe(true)
	})
})

describe('geometry', () => {
	it('keeps a preset’s adjust values as the formulas the source stated', async () => {
		const card = (await nodeNamed('autoshape', 'card')) as ShapeNode
		expect(card.geometry).toStrictEqual({
			kind: 'preset',
			preset: 'roundRect',
			adjustValues: { adj: 'val 13333' },
		})
	})

	it('keeps a freeform’s command list in the read model’s own vocabulary', async () => {
		const freeform = (await nodeNamed('custgeom', 'freeform')) as ShapeNode
		expect(freeform.geometry.kind).toBe('custom')
		if (freeform.geometry.kind !== 'custom') return
		const path = freeform.geometry.paths[0]
		expect(path?.commands.map((command) => command.cmd)).toEqual(['moveTo', 'lnTo', 'cubicBezTo', 'close'])
		// Path units, not EMU: the denominators are the path's own `@w`/`@h`.
		expect(path?.w).toBe(2_743_200)
	})
})

describe('media identity is shared with the contract model', () => {
	it('names an asset the same way DeckIr does', async () => {
		// Two names for one image is how a picture stops being traceable through the
		// loop. The join is by content hash, so it holds even though the package's
		// partname and upstream's asset name have nothing in common.
		const entry = CORPUS.find((candidate) => candidate.name === 'picture')
		if (!entry) throw new Error('corpus lost its picture deck')
		const { deck, render } = await importDeck(await corpusBytes(entry))

		const picture = flatten(render.slides[0]?.nodes ?? []).find((node) => node.kind === 'picture')
		expect(picture?.kind).toBe('picture')
		if (picture?.kind !== 'picture') return

		expect(deck.assets.map((asset) => asset.name)).toContain(picture.asset.$asset)
		const manifest = render.assets.find((asset) => asset.name === picture.asset.$asset)
		expect(manifest?.contentType).toBe('image/png')
		expect(manifest?.sha256).toMatch(/^[0-9a-f]{64}$/)
		// The manifest carries identity, never bytes — the island is JSON.
		expect(Object.keys(manifest ?? {})).toEqual(['name', 'contentType', 'byteLength', 'sha256'])
	})
})

describe('links resolve to identities, not part names', () => {
	it('turns an internal jump into a slide number and keeps an external url', async () => {
		const ir = await importCorpus('hyperlink')
		const runs = ir.slides.flatMap((slide) =>
			flatten(slide.nodes).flatMap((node) =>
				node.kind === 'shape' ? (node.text?.paragraphs.flatMap((paragraph) => paragraph.runs) ?? []) : []
			)
		)
		const jump = runs.find((run) => run.props.hyperlink?.targetSlide !== undefined)
		expect(jump?.props.hyperlink).toStrictEqual({ url: null, targetSlide: 2, tooltip: 'Appendix' })

		const external = runs.find((run) => run.props.hyperlink?.url !== null && run.props.hyperlink !== undefined)
		expect(external?.props.hyperlink?.url).toBe('https://example.invalid/spec')
	})
})

describe('fidelity notes arrive attached to the slide that produced them', () => {
	it('carries upstream’s notes through rather than re-deriving them', async () => {
		const ir = await importCorpus('chart')
		const constructs = ir.slides[0]?.fidelity.map((note) => note.construct) ?? []
		expect(constructs).toContain('chart.workbook')
		// Scoped to the slide, addressed by shape name — the model has no per-node
		// note list, and the note already carries the address.
		expect(ir.slides[0]?.fidelity.every((note) => note.slideNumber === 1)).toBe(true)
	})

	it('draws a chart as a placeholder while leaving its slide authored', async () => {
		// `render: 'placeholder'` is a rendering fact and `source: 'carried'` is a
		// fidelity one. A chart is the case that separates them: nothing can draw it
		// in a browser, and it round-trips through `addChart` perfectly.
		const ir = await importCorpus('chart')
		const chart = ir.slides[0]?.nodes[0]
		expect(chart?.kind).toBe('opaque')
		expect(chart?.render).toBe('placeholder')
		expect(ir.slides[0]?.source).toBe('authored')
		expect(ir.slides[0]?.residual).toBeNull()
	})
})

describe('tables', () => {
	it('addresses cells by table id, row and column', async () => {
		const grid = await nodeNamed('table', 'grid')
		expect(grid.kind).toBe('table')
		if (grid.kind !== 'table') return
		expect(grid.rows[0]?.cells[0]?.id).toBe(`${grid.id}.r0c0`)
		expect(grid.columns).toHaveLength(2)
		expect(grid.rows).toHaveLength(3)
	})

	it('imports per-cell paint rather than a table-wide approximation', async () => {
		const grid = await nodeNamed('table', 'grid')
		if (grid.kind !== 'table') return
		const header = grid.rows[0]?.cells[0]
		const bodyCell = grid.rows[1]?.cells[0]
		expect(header?.fill).toStrictEqual({ kind: 'solid', color: { kind: 'srgb', hex: '250F6B' } })
		expect(header?.borders.top).toStrictEqual({
			kind: 'line',
			widthPt: 1,
			color: { kind: 'srgb', hex: 'CCCCCC' },
			dash: 'solid',
		})
		expect(bodyCell?.fill).not.toStrictEqual(header?.fill)
	})

	it('distinguishes a cell that suppresses its fill from one that states none', async () => {
		// The cell-side twin of `distinguishes an explicit no-fill from an unstated one`
		// above, and the same flattening it exists to prevent: before
		// `TableCell.fillNoFill` (ts-pptx 3.1.0, upstream #7) every colour accessor
		// reported `null` for both cells here, so a suppressed cell was reported as
		// inheriting the table style's shading — and a renderer painting that model
		// fills a cell the deck asked to see through.
		const cells = await nodeNamed('cell-fill', 'cells')
		if (cells.kind !== 'table') return
		expect(cells.rows[1]?.cells[1]?.fill).toStrictEqual({ kind: 'none' })
		expect(cells.rows[1]?.cells[0]?.fill).toStrictEqual({ kind: 'inherit' })
	})
})
