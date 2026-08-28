/**
 * SHA-256, in the two shapes this package hashes things in.
 *
 * One module rather than a copy per layer, and the argument is the one
 * `src/import/assets.ts` already made for exporting the bytes variant: it
 * *defines* {@link import('./ir/render').AssetManifestEntry.sha256}, and the
 * return path verifies resolved bytes against that field. Two implementations of
 * "the manifest hash" is two things that must agree and nothing that makes them,
 * and the failure would be a document reporting every asset as tampered with.
 *
 * The same holds one layer over for the text variant, which produces the
 * `modelHash` and `surfaceHash` the parser compares a document against.
 *
 * `crypto.subtle` is present in Node and in the browser, which is what lets both
 * lanes share this.
 */

/** A digest as lowercase hex. The one place the encoding is written down. */
function hexOf(digest: ArrayBuffer): string {
	return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('')
}

/** Lowercase hex SHA-256 of raw bytes. */
export async function sha256OfBytes(bytes: Uint8Array): Promise<string> {
	// A fresh copy: `digest` wants an ArrayBuffer, and a part's `bytes` may be a
	// view onto a larger buffer, whose tail would otherwise be hashed too.
	return hexOf(await crypto.subtle.digest('SHA-256', Uint8Array.from(bytes)))
}

/** Lowercase hex SHA-256 of a string's UTF-8 bytes. */
export async function sha256Hex(text: string): Promise<string> {
	return hexOf(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text)))
}
