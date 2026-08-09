/**
 * {@link RenderIr} → one self-contained HTML document.
 *
 * The document has two channels and they are not peers:
 *
 * ```
 * IR ──┬──► visual DOM (SVG + HTML)   what a human sees and edits
 *      └──► IR island (JSON)          what the machine reads back
 * ```
 *
 * The visual channel is allowed to be lossy — it is a picture, and several
 * things in it are documented approximations. The island is not lossy at all,
 * and it is the only channel the return path trusts. A slide that renders
 * *wrong* still round-trips *exactly*, which is the property this split exists
 * to buy, and the oracle tests it by deliberately breaking a renderer.
 *
 * ## The renderer never touches the IR
 *
 * Every function reachable from here reads the model and returns strings.
 * Nothing normalizes a colour, resolves a font or clamps a box back into the
 * object it was given, because a renderer that edits the model breaks Invariant
 * R in a way that surfaces much later and somewhere else. This is asserted at
 * test time against a pre-render clone rather than enforced with `Object.freeze`
 * — freezing proves it for one call, the assertion proves it for the model the
 * document was actually built from.
 */

import type { RenderIr, RenderSlide } from '../ir/render'
import { type AssetSource, hydrationScript, renderAssetBlock } from './assets'
import { escapeForScript, INTEGRITY_ID, integrityOf, type Integrity, islandTextOf, ISLAND_ID } from './island'
import { renderNode, type NodeContext } from './node'
import { Defs, fillPaint } from './paint'
import { escapeText } from './text'

/**
 * Where the asset bytes go.
 *
 * A union rather than two optional fields, so `'inline'` cannot be asked for
 * without somewhere to get bytes from — the failure it prevents is a document
 * that claims to be self-contained and silently is not.
 */
export type RenderOptions = { assets: 'ref' } | { assets?: 'inline'; bytes: AssetSource }

export interface RenderedDeck {
	html: string
	/** The two hashes, also embedded in the document. */
	integrity: Integrity
	/** Documented shortfalls — missing bytes, an oversized inline block, an unplaceable node. */
	warnings: string[]
}

/** Slide chrome and the fallback chain the deck's own faces sit in front of. */
const STYLESHEET = `
:root { color-scheme: light }
body { margin: 0; background: #eceef2; font-family: system-ui, sans-serif }
.pxh-deck { display: flex; flex-direction: column; align-items: center; gap: 24px; padding: 24px }
.pxh-slide { background: #fff; box-shadow: 0 1px 6px rgba(0,0,0,.28); width: min(100%, 1280px) }
.pxh-slide > svg { display: block; width: 100%; height: auto }
.pxh-aside { width: min(100%, 1280px); font-size: 13px; color: #454a57 }
.pxh-aside h2 { font-size: 13px; margin: 12px 0 4px }
.pxh-aside ul { margin: 0; padding-left: 18px }
.pxh-carried { color: #7a5200 }
/* A run's own face is written inline; this is only what it falls back to. The
   island keeps the authored face regardless of what the browser painted. */
.pxh-text { font-family: Calibri, Carlito, "Segoe UI", system-ui, sans-serif }
`

function renderSlide(slide: RenderSlide, size: RenderIr['size'], warnings: string[]): string {
	const defs = new Defs()
	const context: NodeContext = { defs, skipped: [] }

	// Order is paint order: the background first, then nodes front-to-back in
	// document order, which is the order the array is already in.
	const background = fillPaint(slide.background.fill, defs)
	const nodes = slide.nodes.map((node) => renderNode(node, context)).join('')

	for (const id of context.skipped) {
		warnings.push(`slide ${slide.number}: node ${id} has no resolvable placement and was not drawn`)
	}

	const svg =
		`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${size.w} ${size.h}" ` +
		`role="img" aria-label="Slide ${slide.number}">` +
		defs.render() +
		`<rect x="0" y="0" width="${size.w}" height="${size.h}" ${background.attrs}` +
		`${background.approx === undefined ? '' : ` data-pxh-approx="${background.approx}"`}/>` +
		nodes +
		`</svg>`

	const asides: string[] = []
	if (slide.source === 'carried') {
		asides.push(
			`<p class="pxh-carried">This slide holds a construct the write API cannot author, so it is carried ` +
				`from the source package verbatim rather than transcribed.</p>`
		)
	}
	if (slide.notes !== null && slide.notes !== '') {
		asides.push(`<h2>Notes</h2><p>${escapeText(slide.notes)}</p>`)
	}
	if (slide.fidelity.length > 0) {
		// Surfaced in the page rather than only in the island: these are the losses
		// the deck *declares*, and a preview that hides them reads as a perfect copy.
		const items = slide.fidelity
			.map((note) => `<li><code>${escapeText(note.construct)}</code> — ${escapeText(note.detail)}</li>`)
			.join('')
		asides.push(`<h2>Declared differences</h2><ul>${items}</ul>`)
	}

	return (
		`<section class="pxh-slide" data-pxh-slide="${slide.number}"${slide.hidden ? ' data-pxh-hidden="true"' : ''}>` +
		`${svg}</section>` +
		(asides.length === 0 ? '' : `<aside class="pxh-aside">${asides.join('')}</aside>`)
	)
}

/**
 * Render a deck.
 *
 * Async because both hashes go through `crypto.subtle`, which is the one digest
 * available in Node and the browser alike — the same reason the import path is
 * async.
 */
export async function renderDeck(ir: RenderIr, options: RenderOptions): Promise<RenderedDeck> {
	const warnings: string[] = []
	const body = ir.slides.map((slide) => renderSlide(slide, ir.size, warnings)).join('')

	// The embedded form, not the raw JSON: `modelHash` covers the block's text
	// exactly as written, so the reader can verify it without an inverse escape
	// that does not exist. See `islandTextOf`.
	const island = islandTextOf(ir)
	const integrity = await integrityOf(ir, island)

	// The island is serialized from the model as given, and the asset block from
	// bytes the caller supplies, so neither depends on the other — which is what
	// makes `modelHash` identical in both asset modes.
	const assets = options.assets === 'ref' ? { html: '', warnings: [] } : renderAssetBlock(ir.assets, options.bytes)
	warnings.push(...assets.warnings)

	const html =
		`<!doctype html>\n<html lang="en"><head><meta charset="utf-8"/>` +
		`<meta name="viewport" content="width=device-width,initial-scale=1"/>` +
		`<title>Deck — ${ir.slides.length} slide${ir.slides.length === 1 ? '' : 's'}</title>` +
		`<style>${STYLESHEET}</style></head><body>` +
		`<div class="pxh-deck">${body}</div>` +
		`<script type="application/json" id="${ISLAND_ID}">${island}</script>` +
		`<script type="application/json" id="${INTEGRITY_ID}">${escapeForScript(JSON.stringify(integrity))}</script>` +
		assets.html +
		hydrationScript() +
		`</body></html>`

	return { html, integrity, warnings }
}
