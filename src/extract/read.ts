// @ts-nocheck
/**
 * Bridge from the host window into the slide iframe: serialise `browserExtractor`
 * and `eval` it inside the frame's realm, returning the extracted IR slide model.
 *
 * The `win.eval(fn.toString())` indirection is intentional: the extractor must run
 * against the iframe's `document`/`getComputedStyle`, not the host's.
 */

import { browserExtractor } from './extractor'

export async function readSlideModel(frame, slideSize, fontConfig) {
	const win = frame.iframe.contentWindow
	return win.eval('(' + browserExtractor.toString() + ')')({
		slideWidthIn: slideSize.width,
		slideHeightIn: slideSize.height,
		fonts: fontConfig
	})
}
