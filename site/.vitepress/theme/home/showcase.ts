/**
 * Two legs of the loop, run on the landing page: `importDeck → renderDeck`.
 *
 * The playground runs all four because it makes a claim about the return path.
 * This page makes the smaller one — *the pictures moving past you were drawn by
 * the library, here, now* — so it stops where the picture is produced. Framework
 * free for the same reason `playground/loop.ts` is: this is `pptx-html`'s public
 * API used the way a consumer would use it, and it is the part worth
 * typechecking. The Vue side owns pixels and nothing else.
 *
 * ## Why the document is cut into slides
 *
 * `renderDeck` writes one document holding every slide, which is right for a
 * preview and wrong for a marquee — the row needs each slide as its own card so
 * it can be placed, scaled and moved independently. So the document is parsed and
 * its `<section class="pxh-slide">` elements are taken out **verbatim**, along
 * with the renderer's own `<style>`. Nothing is re-written on the way past: what
 * a card shows is the markup `renderDeck` produced, and the day the renderer
 * draws something differently the marquee changes with it.
 *
 * The stylesheet travels with the slides because it has to. A slide's own faces
 * are written inline, but the fallback chain, the white ground and the `width:
 * 100%` on the SVG live in that sheet, and a card rendered without it is not the
 * renderer's output any more.
 */

import { importDeck, renderDeck } from 'pptx-html'
import { SHOWCASE, type ShowcaseDeck } from './decks.ts'

export interface ShowcaseSlide {
	/** `<deck>-<n>`, unique across both rows — the marquee's key. */
	key: string
	number: number
	/** The `<section class="pxh-slide">` element, exactly as the renderer wrote it. */
	markup: string
}

export interface ShowcaseRow {
	name: string
	title: string
	blurb: string
	slides: ShowcaseSlide[]
	/** The renderer's own stylesheet, to be installed beside each slide. */
	stylesheet: string
	/** What `renderDeck` could not draw faithfully. Reported, never swallowed. */
	warnings: string[]
	/**
	 * The `.pptx` the row was drawn from — the same bytes `importDeck` was handed,
	 * kept so the row can offer them for download.
	 *
	 * It is the writer's output and not a re-emit: a visitor opening this file in
	 * PowerPoint is looking at the input to the pictures beside it, which is the
	 * only version of that comparison worth offering. Round-tripping is the
	 * playground's claim and needs the other two legs to mean anything.
	 */
	source: Uint8Array
	/** What to call `source` on disk. */
	file: string
}

/**
 * Build, import and render one deck, and hand back its slides as cards.
 *
 * `assets: 'ref'` rather than `'inline'`: these decks carry no media at all, so
 * there are no bytes to embed, and asking for the inline mode would mean handing
 * in an asset source that is never consulted.
 */
async function renderRow(deck: ShowcaseDeck): Promise<ShowcaseRow> {
	const source = await deck.build()
	const imported = await importDeck(source)
	const rendered = await renderDeck(imported.render, { assets: 'ref' })

	const document = new DOMParser().parseFromString(rendered.html, 'text/html')
	const stylesheet = [...document.querySelectorAll('style')].map((style) => style.textContent ?? '').join('\n')
	const slides = [...document.querySelectorAll('section.pxh-slide')].map((section, index) => ({
		key: `${deck.name}-${index + 1}`,
		number: index + 1,
		markup: section.outerHTML,
	}))

	if (slides.length === 0) throw new Error(`${deck.name} rendered no slides`)
	return {
		name: deck.name,
		title: deck.title,
		blurb: deck.blurb,
		slides,
		stylesheet,
		warnings: rendered.warnings,
		source,
		file: deck.file,
	}
}

/** Both rows, rendered in parallel — neither depends on the other. */
export async function showcase(): Promise<ShowcaseRow[]> {
	return await Promise.all(SHOWCASE.map(renderRow))
}

/**
 * The document a single card holds, for installing in a shadow root.
 *
 * A shadow root rather than an iframe, and rather than plain page markup. The
 * page's own stylesheet must not reach a slide — a preview restyled by VitePress
 * is a fake preview, which is the same reason the playground puts its preview in
 * a frame — but thirty-two frames scrolling at sixty hertz is a different
 * problem, and a shadow root buys the isolation without the documents.
 *
 * `all: initial` is what closes the last gap: style rules do not cross a shadow
 * boundary but *inheritance* does, and the slide's text lives in a
 * `<foreignObject>` full of HTML that would happily inherit VitePress's font and
 * colour. Resetting every inherited property on the host means the only thing
 * styling that HTML is the renderer's sheet and the faces written inline.
 */
export function cardStyles(stylesheet: string): string {
	return `${stylesheet}
:host { all: initial; display: block }
/* The card's own frame carries the elevation, so the slide inside is flat and
   fills it — the renderer sizes for a scrolling preview, not for a tile. */
.pxh-slide { width: 100%; box-shadow: none }
/* Notes, carried-slide notices and declared differences are the playground's
   job: they are prose beside a slide, and this row shows slide faces. The
   caption under the rows says where to read them. */
.pxh-aside { display: none }
`
}
