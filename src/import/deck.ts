/**
 * `.pptx` → the two IRs.
 *
 * ```
 *                        ┌─ readModelToIr ──► DeckIr    (the contract: emit + diff)
 * bytes ──► Presentation ┤
 *                        └─ this file ─────► RenderIr   (the paint model: render)
 * ```
 *
 * Both models are built from **one loaded `Presentation`**, and that is the whole
 * structural guarantee: two independent loads could disagree about the source
 * deck, and then a difference between them would be indistinguishable from a
 * difference the loop introduced. (They are not one literal walk — upstream's
 * mapper owns its own traversal — so this is the property that is actually
 * available, and it is the one that matters.)
 *
 * The contract side is **not reimplemented here**. `readModelToIr` is called, its
 * assets are the media identity both models share, and its `fidelity` notes are
 * merged into the matching slides so a renderer can label what was lost without
 * re-deriving anything.
 *
 * This layer never touches XML. Everything comes through `@shbernal/ts-pptx/read`'s
 * typed object graph, which is what makes it browser-capable — the same code runs
 * in Chromium, where the point of the loop lives.
 */

import {
	type BackgroundFill,
	type OpcInput,
	Presentation,
	type Slide,
	type SlideBackground,
} from '@shbernal/ts-pptx/read'
import { type DeckIr, type FidelityNote, readModelToIr } from '@shbernal/ts-pptx/script'
import { type Background, type Fill, IR_VERSION, type RenderIr, type RenderSlide } from '../ir/render'
import { type AssetIndex, buildAssetIndex } from './assets'
import type { ImportScope } from './context'
import { colorOf, gradientOf, pictureFillOf } from './paint'
import { residualOf } from './residual'
import { nodeOf } from './shape'

/** Both models, plus the package they were read from. */
export interface ImportedDeck {
	/** The loaded package. Kept because the emit lane needs it as its template. */
	presentation: Presentation
	/** The round-trip contract, upstream's. Emit writes this and `diffDeckIr` judges it. */
	deck: DeckIr
	/** The paint model, this package's. */
	render: RenderIr
	/** Bytes behind the render manifest's references, for a caller that must write them out. */
	assets: AssetIndex
}

/** Load a deck and import it. Accepts everything `Presentation.load` does. */
export async function importDeck(input: OpcInput): Promise<ImportedDeck> {
	return importPresentation(await Presentation.load(input))
}

export async function importPresentation(presentation: Presentation): Promise<ImportedDeck> {
	const deck = readModelToIr(presentation)
	const assets = await buildAssetIndex(presentation.opc, deck)

	const slideNumberByPart = new Map<string, number>()
	for (const slide of presentation.slides) slideNumberByPart.set(slide.part.partName, slide.index + 1)

	const size = presentation.slideSize
	const render: RenderIr = {
		irVersion: IR_VERSION,
		size: { w: size?.widthEmu ?? 0, h: size?.heightEmu ?? 0 },
		slides: presentation.slides.map((slide) => renderSlideOf(slide, deck, assets, slideNumberByPart)),
		assets: assets.manifest,
	}

	return { presentation, deck, render, assets }
}

function renderSlideOf(
	slide: Slide,
	deck: DeckIr,
	assets: AssetIndex,
	slideNumberByPart: ReadonlyMap<string, number>
): RenderSlide {
	const number = slide.index + 1
	const contract = deck.slides.find((entry) => entry.number === number)

	// Upstream's notes for this slide come first, so a note the *importer* files
	// about the same shape reads as an addition to them rather than a competing
	// account. Both are in one vocabulary; there is no second taxonomy here.
	const notes: FidelityNote[] = deck.fidelity.filter((entry) => entry.slideNumber === number)

	const scope: ImportScope = { slideNumber: number, shapeName: null, notes, assets, slideNumberByPart }
	const nodes = slide.shapes.map((shape) => nodeOf(shape, scope))

	const source = contract?.source ?? 'authored'
	return {
		number,
		source,
		layout: contract?.layout ?? null,
		hidden: slide.hidden,
		background: backgroundOf(slide.background, scope),
		nodes,
		// `notesText` is `''` for a slide with no notes part at all, which is not the
		// same as a notes part holding nothing. The presence of the part is the
		// distinction worth keeping, so `null` means there is none.
		notes: slide.notesSlide === null ? null : slide.notesText,
		// Only a carried slide gets one: a residual beside an authored slide is a
		// second copy of the same content that nothing keeps in step, and it would
		// double the size of every island for nothing.
		residual: source === 'carried' ? residualOf(slide, assets) : null,
		fidelity: notes,
	}
}

/**
 * A slide's effective background.
 *
 * `themeRef` — a `p:bgRef` naming an entry in the theme's style matrix, and the
 * common case in a generated deck — has no representation of its own here: it is
 * resolved through `resolvedFill` into the concrete fill it stands for. That is a
 * paint decision rather than a flattening one. The reference itself is `DeckIr`'s
 * to carry, and it does; nothing emits from this.
 */
function backgroundOf(background: SlideBackground | null, scope: ImportScope): Background {
	if (background === null) return { source: 'master', fill: { kind: 'inherit' } }
	if (background.type !== 'themeRef') return { source: background.source, fill: fillOfBackground(background, scope) }

	const resolved = background.resolvedFill
	if (resolved !== null) return { source: background.source, fill: fillOfBackground(resolved, scope) }
	const color = colorOf(null, background.color)
	return { source: background.source, fill: color === undefined ? { kind: 'inherit' } : { kind: 'solid', color } }
}

function fillOfBackground(background: BackgroundFill, scope: ImportScope): Fill {
	switch (background.type) {
		case 'none':
			return { kind: 'none' }
		case 'solid': {
			const color = colorOf(null, background.color)
			return color === undefined ? { kind: 'inherit' } : { kind: 'solid', color }
		}
		case 'gradient': {
			const gradient = gradientOf(background.gradient, scope)
			return gradient === null ? { kind: 'inherit' } : { kind: 'gradient', gradient }
		}
		case 'pattern':
			return {
				kind: 'pattern',
				preset: background.preset,
				foreground: colorOf(null, background.foreground) ?? null,
				background: colorOf(null, background.background) ?? null,
			}
		case 'image':
			return pictureFillOf(background.picture, scope)
	}
}
