/**
 * Base64, both directions, without Node's `Buffer`.
 *
 * `btoa`/`atob` are the pair present in both runtimes, which is what lets the
 * loop and the heuristic lane share this. `Uint8Array.fromBase64` / `.toBase64`
 * would replace both and are **not** available on Node 24.16, so the manual
 * codecs stay until the floor moves.
 *
 * **There is a third copy, deliberately.** `hydrationScript()` in
 * `src/render/assets.ts` emits a decoder as *source text* that runs inside the
 * rendered document, where it cannot import anything. A change to
 * {@link bytesOfBase64} has a mirror there.
 */

/**
 * Base64 for bytes.
 *
 * `btoa` takes a string of code units below 256. The chunking is not an
 * optimisation: spreading a multi-megabyte array into `String.fromCharCode(...)`
 * in one call overflows the argument-list limit and throws, which would make
 * large images fail exactly where they matter most.
 */
export function base64Of(bytes: Uint8Array): string {
	const CHUNK = 0x8000
	let binary = ''
	for (let offset = 0; offset < bytes.length; offset += CHUNK) {
		binary += String.fromCharCode(...bytes.subarray(offset, offset + CHUNK))
	}
	return btoa(binary)
}

/** Bytes for base64. The mirror of {@link base64Of}. */
export function bytesOfBase64(encoded: string): Uint8Array {
	const binary = atob(encoded)
	const bytes = new Uint8Array(binary.length)
	for (let index = 0; index < binary.length; index++) bytes[index] = binary.charCodeAt(index)
	return bytes
}
