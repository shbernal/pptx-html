/**
 * Handing a `.pptx` to the visitor, from bytes that are already in the tab.
 *
 * Shared by the playground and the home page because both are making the same
 * point with it: the file you get is the one the library produced here, not a
 * prepared artefact fetched from the server. There is no network in this
 * function, and there is deliberately nowhere for one to be added — an object
 * URL over bytes held in memory is the whole mechanism.
 */

/** The OPC media type for a presentation. PowerPoint keys off it, not the extension. */
const PPTX = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'

export function download(bytes: Uint8Array, name: string): void {
	// `slice()` rather than the view itself: a `Uint8Array` may be backed by a
	// `SharedArrayBuffer`, which is not a `BlobPart`, and the copy is the honest
	// way to narrow it. One click, one deck-sized copy — not a cost worth a cast.
	const blob = new Blob([bytes.slice()], { type: PPTX })
	const url = URL.createObjectURL(blob)
	const anchor = document.createElement('a')
	anchor.href = url
	anchor.download = name
	anchor.click()
	URL.revokeObjectURL(url)
}
