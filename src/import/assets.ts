/**
 * Media identity, shared between the two models.
 *
 * `readModelToIr` already hash-addresses every image it carries into
 * `DeckIr.assets` and refers to it with the `$asset` sigil, so the render model
 * must use *those* names or the loop ends up with two identities for one image —
 * and then nothing can tell whether the picture that came back is the picture
 * that went in.
 *
 * The join is by **content hash**, not by partname, because upstream's asset
 * names (`image1.png`) are its own and the package's partnames
 * (`/ppt/media/image-1-1.png`) are the source's, and nothing publishes the map
 * between them. Hashing both sides is the only join that does not depend on an
 * internal naming convention holding still — and content addressing is what the
 * IR claims media identity *is*, so this is the rule applied rather than a
 * workaround for a missing accessor.
 */

import type { AssetIr, AssetRef, DeckIr } from '@shbernal/ts-pptx/script'
import type { OpcPackage } from '@shbernal/ts-pptx/read'
import type { AssetManifestEntry } from '../ir/render'

export interface AssetIndex {
	/** Every asset any slide references. Bytes are never here — see `ir/render`. */
	manifest: AssetManifestEntry[]
	/** The reference for a package partname, or `null` when it holds no media we carry. */
	refFor(partName: string | null): AssetRef | null
	/** The bytes behind a reference, for a caller that has to write them out. */
	bytesFor(ref: AssetRef): Uint8Array | undefined
}

/** Lowercase hex SHA-256. `crypto.subtle` is present in Node and in the browser. */
async function sha256Hex(bytes: Uint8Array): Promise<string> {
	// A fresh copy: `digest` wants an ArrayBuffer, and a part's `bytes` may be a
	// view onto a larger buffer, whose tail would otherwise be hashed too.
	const digest = await crypto.subtle.digest('SHA-256', Uint8Array.from(bytes))
	return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('')
}

/** Is this part media a picture node or a picture fill could point at? */
function isMediaPart(contentType: string): boolean {
	return contentType.startsWith('image/')
}

/**
 * A name no other asset has taken.
 *
 * Two parts in different directories can share a filename while holding
 * different bytes; without this they would collapse onto one manifest entry and
 * one of the two pictures would silently render as the other.
 */
function uniqueName(candidate: string, taken: ReadonlyMap<string, unknown>): string {
	if (!taken.has(candidate)) return candidate
	const dot = candidate.lastIndexOf('.')
	const stem = dot > 0 ? candidate.slice(0, dot) : candidate
	const extension = dot > 0 ? candidate.slice(dot) : ''
	for (let index = 2; ; index++) {
		const next = `${stem}-${index}${extension}`
		if (!taken.has(next)) return next
	}
}

/**
 * Build the index for one deck.
 *
 * Every media part in the package is hashed, not only the ones upstream carried:
 * a part upstream *dropped* still has to be drawable, so it is registered under a
 * name derived from its partname. Such a name exists in the render manifest and
 * not in `DeckIr.assets`, which is exactly the asymmetry it should have — the
 * bytes reach the page and do not reach the emitted deck, and the note upstream
 * already filed says so.
 */
export async function buildAssetIndex(opc: OpcPackage, deck: DeckIr): Promise<AssetIndex> {
	const carriedByHash = new Map<string, AssetIr>()
	for (const asset of deck.assets) carriedByHash.set(await sha256Hex(asset.bytes), asset)

	const manifest: AssetManifestEntry[] = []
	const refByPart = new Map<string, AssetRef>()
	const bytesByName = new Map<string, Uint8Array>()
	const nameByHash = new Map<string, string>()

	for (const [partName, part] of opc.parts) {
		if (!isMediaPart(part.contentType)) continue
		const hash = await sha256Hex(part.bytes)
		const seen = nameByHash.get(hash)
		if (seen !== undefined) {
			// Same bytes as a part already indexed: one asset, two references.
			// Deduplication here is what stops `importSlide`'s shared media from
			// being counted twice in the manifest the island carries.
			refByPart.set(partName, { $asset: seen })
			continue
		}
		const carried = carriedByHash.get(hash)
		const name = uniqueName(carried?.name ?? partName.replace(/^.*\//, ''), bytesByName)
		nameByHash.set(hash, name)
		refByPart.set(partName, { $asset: name })
		bytesByName.set(name, part.bytes)
		manifest.push({
			name,
			contentType: carried?.contentType ?? part.contentType,
			byteLength: part.bytes.byteLength,
			sha256: hash,
		})
	}

	return {
		manifest,
		refFor: (partName) => (partName === null ? null : (refByPart.get(partName) ?? null)),
		bytesFor: (ref) => bytesByName.get(ref.$asset),
	}
}
