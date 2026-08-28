/**
 * Getting bytes back for the manifest's {@link AssetRef}s.
 *
 * The island carries names, content types, lengths and hashes; it never carries
 * bytes. So the return path has two places to look, and the order between them is
 * a decision rather than a convenience:
 *
 * 1. an explicit {@link AssetResolver} supplied by the caller, and
 * 2. the document's own inline asset block.
 *
 * **The explicit resolver wins.** A caller that supplies one has a source of
 * truth the document does not — a media store, the original package — and the
 * inline block is by construction a copy that travelled through a browser. When
 * the two disagree, the copy is the one to doubt.
 *
 * ## Every byte is checked, and a failure here is not a lane
 *
 * `AssetManifestEntry.sha256` exists for this moment. Bytes that do not match it
 * are not a degraded picture, they are a different file, and pasting a different
 * file into the emitted deck under the original's name is exactly the silent loss
 * the charter forbids. So a hash mismatch **throws**.
 *
 * An asset that resolves to *nothing* is different, and is a warning: the model
 * still knows the asset exists, the manifest still states what it was, and a
 * caller that renders with `assets: 'ref'` and then parses without a resolver has
 * done nothing wrong. What it gets back is a deck that cannot re-embed that
 * image, and it is told so by name.
 */

import { bytesOfBase64 } from '../base64'
import { sha256OfBytes } from '../hash'
import type { AssetManifestEntry } from '../ir/render'

/**
 * Where a caller says bytes come from.
 *
 * It takes the whole manifest entry rather than just the name so a resolver can
 * key on the content hash — the identity media actually has in this model — and
 * not on a package partname that means nothing outside the source deck.
 */
export type AssetResolver = (entry: AssetManifestEntry) => Promise<Uint8Array | undefined> | Uint8Array | undefined

export interface ResolvedAssets {
	/** Bytes for a manifest name, or `undefined` when nothing supplied them. */
	bytesFor(name: string): Uint8Array | undefined
	/** Manifest entries nothing could supply, by name. */
	missing: string[]
	warnings: string[]
}

/**
 * A resolver over the document's own inline block, or `null` when there is none.
 *
 * Exported so a caller can compose it — resolve from the document, fall back to a
 * store, or the reverse — rather than being limited to the precedence
 * {@link resolveAssets} applies.
 */
export function inlineAssetResolver(blockText: string | null): AssetResolver | null {
	if (blockText === null) return null
	const encoded = JSON.parse(blockText) as Record<string, string>
	return (entry) => {
		const data = encoded[entry.name]
		return data === undefined ? undefined : bytesOfBase64(data)
	}
}

/**
 * One manifest entry, resolved or not. A union rather than a nullable digest so
 * that "these bytes arrived" and "they hash to this" are one fact: there is no
 * state where a resolver returned bytes and the digest is missing.
 */
type Resolution =
	| { entry: AssetManifestEntry; resolved: undefined }
	| { entry: AssetManifestEntry; resolved: Uint8Array; digest: string }

export interface ResolveOptions {
	/** The caller's own source of bytes. Consulted first. */
	resolver?: AssetResolver
	/** The document's inline block, as text. Consulted when the resolver has nothing. */
	inline?: string | null
}

/**
 * Resolve the whole manifest, verifying as it goes.
 *
 * The length check before the hash is not an optimisation — it produces a much
 * better message for the overwhelmingly common cause of a mismatch, which is a
 * resolver keyed on the wrong name and returning some *other* asset.
 */
export async function resolveAssets(
	manifest: readonly AssetManifestEntry[],
	options: ResolveOptions = {}
): Promise<ResolvedAssets> {
	const inline = inlineAssetResolver(options.inline ?? null)
	const bytes = new Map<string, Uint8Array>()
	const missing: string[] = []
	const warnings: string[] = []

	// Resolve and hash every entry at once, then *decide* in manifest order.
	// Nothing about resolving one entry depends on another, but the verdict does:
	// a caller whose deck has two bad assets should be told about the first one in
	// the manifest, every run, rather than about whichever resolver happened to
	// settle first.
	const resolutions = await Promise.all(
		manifest.map(async (entry): Promise<Resolution> => {
			const resolved = (await options.resolver?.(entry)) ?? (await inline?.(entry))
			if (resolved === undefined) return { entry, resolved: undefined }
			return { entry, resolved, digest: await sha256OfBytes(resolved) }
		})
	)

	for (const resolution of resolutions) {
		const entry = resolution.entry
		if (resolution.resolved === undefined) {
			missing.push(entry.name)
			warnings.push(
				`asset ${JSON.stringify(entry.name)} (${entry.contentType}, ${entry.byteLength} bytes) could not be resolved; the deck will be emitted without it`
			)
			continue
		}

		const { resolved, digest } = resolution
		if (resolved.byteLength !== entry.byteLength) {
			throw new Error(
				`asset ${JSON.stringify(entry.name)} resolved to ${resolved.byteLength} bytes and the manifest states ${entry.byteLength}; this is a different file, not a damaged one`
			)
		}
		if (digest !== entry.sha256) {
			throw new Error(
				`asset ${JSON.stringify(entry.name)} does not match its manifest hash (${entry.sha256.slice(0, 12)}… vs ${digest.slice(0, 12)}…)`
			)
		}
		bytes.set(entry.name, resolved)
	}

	return { bytesFor: (name) => bytes.get(name), missing, warnings }
}
