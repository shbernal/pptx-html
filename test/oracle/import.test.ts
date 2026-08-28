/**
 * What the importer must get right, checked against real decks rather than
 * hand-written input — because the mapping's failure mode is not "throws", it is
 * "produces something plausible that the source did not say".
 *
 * Lives under `test/oracle/` rather than `test/unit/` for one reason: every case
 * here needs a corpus deck built by the writer, which is the same generation the
 * oracle project owns. No DOM is involved.
 */

import { LAYOUT_NOTE_PREFIX } from '@shbernal/ts-pptx/script'
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
			// The JSON round trip is the assertion, not a way to copy: `structuredClone`
			// here would prove that the model survives structured cloning, which is not
			// the claim this test makes.
			// oxlint-disable-next-line unicorn/prefer-structured-clone
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
		// The cast is the honest type. `constructor` is declared as always present,
		// and the one object that has none is `Object.create(null)` — which is inside
		// this branch, because a null prototype is not `Object.prototype`.
		const ctor = (value as { constructor?: { name?: string } }).constructor
		offenders.push(`${path}: ${ctor?.name ?? 'exotic object'}`)
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

	it('keeps a gradient stop’s transform list rather than the hex it flattens to', async () => {
		// Before ts-pptx 3.6.0 (upstream #26) a stop reported a token and a painted hex
		// and nothing between them, so this list was empty for every stop in the corpus
		// — and an empty list is the same shape as "the deck stated no transform". The
		// two are only distinguishable once one stop has a transform and its neighbour
		// does not.
		const faded = (await nodeNamed('color-transform', 'faded')) as ShapeNode
		if (faded.fill.kind !== 'gradient') throw new Error('the faded shape lost its gradient')
		expect(faded.fill.gradient.stops[0]?.color).toStrictEqual({
			kind: 'scheme',
			slot: 'accent1',
			transforms: [{ name: 'alpha', value: '60000' }],
			effectiveHex: '250F6B',
			alpha: 0.6,
		})
		expect(faded.fill.gradient.stops[1]?.color).toStrictEqual({
			kind: 'scheme',
			slot: 'accent2',
			transforms: [],
			effectiveHex: '9B1B30',
		})
	})

	it('keeps a cell edge’s opacity, which the flattened colour dropped outright', async () => {
		// The other half of upstream #26, one construct along and a worse loss: a
		// `CellBorder` gave a resolved hex with no alpha beside it, so a 35 %
		// transparent rule was imported fully opaque and drawn as one.
		const ruled = await nodeNamed('color-transform', 'ruled')
		if (ruled.kind !== 'table') throw new Error('the ruled node lost its table')
		expect(ruled.rows[0]?.cells[0]?.borders.top).toStrictEqual({
			kind: 'line',
			widthPt: 2,
			color: { kind: 'srgb', hex: '250F6B', alpha: 0.65 },
			dash: 'solid',
		})
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

	it('resolves an italic the run inherits without writing it into props', async () => {
		// `resolvedItalic` arrived in ts-pptx 3.6.0 (upstream #27). Until then italic
		// was the one property on the run with no resolved counterpart, so a run that
		// takes its slant from the master was painted upright — and because `props`
		// stayed empty either way, the deck round-tripped while the preview lied.
		const body = (await nodeNamed('layout-placeholder', 'body')) as ShapeNode
		const run = body.text?.paragraphs[0]?.runs[0]
		expect(run?.props).toStrictEqual({})
		expect(run?.resolved.italic).toBe(true)
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

describe('the template’s furniture arrives as chrome, not as nodes', () => {
	it('reaches a layout’s non-placeholder shapes and keeps them off the slide’s own tree', async () => {
		// The gap this closes was the largest single one measured: `slide.shapes` is
		// the slide's tree alone, so a template's band and wordmark were absent from
		// every preview. They are here now, and they are *here* — in `chrome` — rather
		// than in `nodes`, which is the whole distinction.
		const ir = await importCorpus('layout-chrome')
		const slide = ir.slides[0]
		if (!slide) throw new Error('layout-chrome lost its slide')

		const chrome = flatten(slide.chrome)
		expect(chrome.map((node) => node.kind)).toEqual(['shape', 'shape'])

		const band = chrome[0]
		expect(band?.kind).toBe('shape')
		if (band?.kind !== 'shape') return
		expect(band.fill).toStrictEqual({ kind: 'solid', color: { kind: 'srgb', hex: '250F6B' } })
		// Slide-absolute EMU, straight off the read model: a 10in × 0.45in band.
		expect(band.placement?.box).toStrictEqual({ x: 0, y: 0, w: 9_144_000, h: 411_480 })

		const wordmark = chrome[1]
		expect(wordmark?.kind === 'shape' && wordmark.text?.paragraphs[0]?.runs[0]?.text).toBe('ACME')

		// The slide's own title is a node, and it is the only one. A layout
		// placeholder that leaked into chrome would show up here as a third shape
		// carrying the prompt text.
		expect(slide.nodes).toHaveLength(1)
		expect(slide.nodes[0]?.name).toBe('title')
		expect(chrome.map((node) => node.name)).not.toContain('title')
	})

	it('namespaces a chrome id by tier so it cannot be addressed as a slide node', async () => {
		// `p:cNvPr/@id` is unique within *a* tree, and chrome comes from two further
		// trees — so without the tier in the id a layout shape and a slide shape
		// collide, and an edit addressed to one would find the other.
		const ir = await importCorpus('layout-chrome')
		const slide = ir.slides[0]
		if (!slide) throw new Error('layout-chrome lost its slide')

		for (const node of flatten(slide.chrome)) expect(node.id).toMatch(/^s1\.(layout|master)\.sp\d+$/)
		const slideIds = new Set(flatten(slide.nodes).map((node) => node.id))
		for (const node of flatten(slide.chrome)) expect(slideIds.has(node.id)).toBe(false)
	})

	it('files no fidelity note of its own, and lets no layout note reach the slide', async () => {
		// Chrome is not emitted — the layout part travels in the template package and
		// the emitted deck binds to it — so a note filed while *walking* it would
		// declare a loss the round trip never takes, addressed to a shape that is not
		// on the slide. `forChrome` replaces the notes array, which is what enforces it.
		const entry = CORPUS.find((candidate) => candidate.name === 'layout-chrome')
		if (!entry) throw new Error('the corpus lost its layout-chrome deck')
		const { deck, render } = await importDeck(await corpusBytes(entry))

		const chromeNames = new Set(flatten(render.slides[0]?.chrome ?? []).map((node) => node.name))
		expect(chromeNames).toEqual(new Set(['Shape 0', 'Text 1']))

		// The second half. A note about a layout must not land on the slide: it would
		// show the reader "Text 1 lost its indent" beneath a slide whose own content is
		// intact, about a shape they cannot select and an edit the return path never
		// makes. The mechanism is that such a note carries **no slide number**, so the
		// `slideNumber === number` filter that builds a slide's notes drops it.
		//
		// This used to be anchored on a live `layout.text.indent`, the last note left on
		// this deck once upstream started re-authoring a layout's furniture into
		// `defineSlideMaster({ objects })`. ts-pptx 3.2.0 retired that note too — it
		// carries a paragraph's own margins now — so the corpus produces no
		// `layout.`-prefixed note at all, and anchoring on one would assert nothing
		// while looking like it asserts something.
		//
		// So the guard is asserted through the *class* a layout note belongs to: notes
		// with no slide number. That class is non-empty here and stays that way — a
		// deck-level loss (theme, master, docProps) is always in it — which is what
		// keeps this from going quietly vacuous a second time.
		expect(deck.fidelity.filter((note) => note.construct.startsWith(LAYOUT_NOTE_PREFIX))).toEqual([])

		const deckLevel = deck.fidelity.filter((note) => note.slideNumber === null)
		expect(deckLevel.length).toBeGreaterThan(0)
		const slideNotes = render.slides[0]?.fidelity ?? []
		for (const note of slideNotes) expect(note.slideNumber).toBe(1)
		for (const note of deckLevel) expect(slideNotes).not.toContain(note)

		for (const note of slideNotes) expect(chromeNames.has(note.shapeName ?? '')).toBe(false)
	})

	it('leaves chrome empty for a deck whose layout carries only placeholders', async () => {
		// The negative case, and it is the common one: `defineSlideMaster` with
		// nothing but placeholders puts no furniture on the layout, and an importer
		// that mistook a placeholder for furniture would draw the prompt text under
		// every slide in the corpus.
		for (const name of ['layout-placeholder', 'master-background', 'text-box']) {
			const ir = await importCorpus(name)
			expect(ir.slides.map((slide) => slide.chrome.length)).toEqual(ir.slides.map(() => 0))
		}
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

	it('points a vector picture at its SVG, not at the raster fallback beside it', async () => {
		// The failure this pins is silent, which is why it is worth a test of its own:
		// an SVG picture resolves to a real part either way, so nothing warns and
		// nothing drops. It just paints the fallback — and upstream writes that
		// fallback as a 1×1 transparent PNG, so a deck of icons renders as a deck of
		// blank boxes that claims to be complete.
		const ir = await importCorpus('picture-svg')
		const glyph = flatten(ir.slides[0]?.nodes ?? []).find((node) => node.kind === 'picture')
		expect(glyph?.kind).toBe('picture')
		if (glyph?.kind !== 'picture') return

		const manifest = ir.assets.find((asset) => asset.name === glyph.asset.$asset)
		expect(manifest?.contentType).toBe('image/svg+xml')

		// And the fallback is still carried, so a consumer that cannot draw SVG has
		// something to fall back *to*. Dropping it would trade one loss for another.
		expect(ir.assets.map((asset) => asset.contentType)).toContain('image/png')
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

	it('reads the two diagonals, which no edge can stand in for', async () => {
		// `a:lnTlToBr` / `a:lnBlToTr`. Both sides of ts-pptx can express one — the
		// read model decodes them and `TableCellProps.diagonal` writes them back — so
		// a cell struck out in the source and plain in the model would be a loss this
		// pipeline chose, which is the one thing `docs/architecture.md` rules out.
		const struck = await nodeNamed('cell-diagonal', 'struck')
		if (struck.kind !== 'table') return
		const [one, both, plain] = struck.rows[0]?.cells ?? []
		expect(one?.borders.tlToBr).toStrictEqual({
			kind: 'line',
			widthPt: 1,
			color: { kind: 'srgb', hex: 'C00000' },
			dash: 'solid',
		})
		// Stated on one corner only: the other diagonal is unstated, not the same line.
		expect(one?.borders.blToTr).toStrictEqual({ kind: 'inherit' })
		expect(both?.borders.blToTr).toStrictEqual({
			kind: 'line',
			widthPt: 0.75,
			color: { kind: 'srgb', hex: '250F6B' },
			dash: 'sysDash',
		})
		expect(plain?.borders.tlToBr).toStrictEqual({ kind: 'inherit' })
		expect(plain?.borders.blToTr).toStrictEqual({ kind: 'inherit' })
	})
})
