/**
 * The asset block and the hydration it enables.
 *
 * Where the bytes live is a **render option**, not a property of the model. That
 * follows upstream's own reasoning for putting `AssetMode` on the print options
 * rather than on `DeckIr`: an IR-level choice would bake one answer into the
 * model and force a re-read to change it. So the island always carries
 * `AssetRef` keys and a manifest, never bytes, and this file decides whether the
 * bytes ride along in the same file.
 *
 * ## The visual channel hydrates; it does not store
 *
 * The obvious implementation — `<img src="data:image/png;base64,…">` on every
 * picture — writes every image into the document twice once the asset block is
 * there too, and more than twice for an image used on several slides. Instead
 * each painted image carries only `data-pxh-asset="<name>"`, and the script
 * below resolves it against the block once, through one Blob URL per asset.
 *
 * The cost is that **the document needs JavaScript to look right**. It does not
 * need JavaScript to round-trip: the island and the asset block are both inert
 * text, and the parser reads them without executing anything. That asymmetry is the
 * whole point, and it costs nothing in reach — a sanitizer aggressive enough to
 * strip `<script>` has already destroyed the island, so there is no context that
 * keeps the round-trip but loses the pictures.
 */

import { base64Of } from '../base64'
import type { AssetManifestEntry } from '../ir/render'
import { ASSETS_ID, escapeForScript, ISLAND_ID } from './island'

/**
 * Where {@link renderAssetBlock} gets bytes for a manifest name.
 *
 * A lookup rather than a map so the caller can stream from whatever it already
 * has — `AssetIndex.bytesFor` from the import path satisfies it directly.
 */
export type AssetSource = (name: string) => Uint8Array | undefined

/** Bytes past which `'ref'` is worth suggesting. Base64 adds a third on top of this. */
export const ASSET_SIZE_WARN_BYTES = 10_000_000

export interface AssetBlock {
	/** The `<script>` element, or `''` under `assets: 'ref'`. */
	html: string
	/** Advice for the caller — an oversized document, or a manifest entry with no bytes. */
	warnings: string[]
}

/**
 * The inline asset block: `{ name: base64 }`, keyed exactly as the manifest and
 * the island's {@link AssetRef}s are.
 *
 * A manifest entry the source cannot supply is a warning rather than an error.
 * The document is still a complete, valid round-trip carrier — the model knows
 * about the asset, the manifest states its hash and length, and only the
 * *picture* is missing. Failing the whole render for a missing preview would
 * trade the thing this project guarantees for the thing it explicitly allows to
 * be lossy.
 */
export function renderAssetBlock(manifest: readonly AssetManifestEntry[], bytes: AssetSource): AssetBlock {
	const warnings: string[] = []
	const encoded: Record<string, string> = {}
	let total = 0

	for (const entry of manifest) {
		const data = bytes(entry.name)
		if (data === undefined) {
			warnings.push(`asset ${JSON.stringify(entry.name)} is in the manifest but no bytes were supplied; it will not be painted`)
			continue
		}
		encoded[entry.name] = base64Of(data)
		total += data.byteLength
	}

	if (total > ASSET_SIZE_WARN_BYTES) {
		warnings.push(
			`${Math.round(total / 1_000_000)} MB of assets are inlined (base64 adds ~33% on top); render with { assets: 'ref' } and resolve them separately if the document is too large`
		)
	}

	// The island's rule, *called* rather than retyped, for the same reason: a `<`
	// inside base64 is impossible, but a manifest *name* comes from a package
	// partname and is not under this file's control. `src/parse/island.ts` scans
	// these blocks with a string scanner, which is safe only because every `<` in
	// every block this package writes is escaped — so the day `escapeForScript`
	// has to disarm something else, this gets it too.
	const json = escapeForScript(JSON.stringify(encoded))
	return { html: `<script type="application/json" id="${ASSETS_ID}">${json}</script>`, warnings }
}

/**
 * The hydration script: manifest name → Blob URL → every element that asked for
 * it by `data-pxh-asset`.
 *
 * One Blob per asset, not per element, so a logo on twelve slides is decoded
 * once. It is emitted as source text rather than a bundled module because the
 * document has to be self-contained — it may be opened from a file:// URL with
 * no server and no import map behind it.
 *
 * It fails quietly and completely: if the asset block is missing (`'ref'` mode,
 * or a sanitizer removed it) every image simply stays blank, which is the
 * correct rendering of "these bytes are not in this file".
 */
export function hydrationScript(): string {
	return `<script>(function(){
var block=document.getElementById(${JSON.stringify(ASSETS_ID)});
if(!block)return;
var assets=JSON.parse(block.textContent||'{}');
var manifest=JSON.parse((document.getElementById(${JSON.stringify(ISLAND_ID)})||{textContent:'{}'}).textContent||'{}').assets||[];
var typeByName={};
manifest.forEach(function(entry){typeByName[entry.name]=entry.contentType});
var urls={};
function urlFor(name){
if(urls[name])return urls[name];
var data=assets[name];
if(!data)return null;
var raw=atob(data),bytes=new Uint8Array(raw.length);
for(var i=0;i<raw.length;i++)bytes[i]=raw.charCodeAt(i);
urls[name]=URL.createObjectURL(new Blob([bytes],{type:typeByName[name]||'application/octet-stream'}));
return urls[name];
}
document.querySelectorAll('[data-pxh-asset]').forEach(function(node){
var url=urlFor(node.getAttribute('data-pxh-asset'));
if(!url)return;
if(node.tagName.toLowerCase()==='image')node.setAttribute('href',url);
else node.setAttribute('src',url);
});
})();</script>`
}
