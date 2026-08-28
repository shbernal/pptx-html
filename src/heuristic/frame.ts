/**
 * Hidden-iframe lifecycle and "settle" waits (fonts, stylesheets, icons) for the
 * heuristic lane. Host-window code, not the stringified extractor payload.
 */

/** The hidden frame a slide is rendered in, and the handle the lane drives it by. */
export interface SlideFrame {
	iframe: HTMLIFrameElement
	load: () => Promise<void>
	resize: (width: number, height: number) => void
	close: () => void
}

/** An `<iconify-icon>` element, narrowed to what the wait loop touches. */
interface IconElement extends Element {
	shadowRoot: ShadowRoot | null
	loadIcon?: () => void
}

export function escapeAttr(value: unknown): string {
	return String(value || '')
		.replace(/&/g, '&amp;')
		.replace(/"/g, '&quot;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
}

export function withBaseHref(html: string): string {
	if (/<base\b/i.test(html)) return html
	const href = (document.baseURI || window.location.href || '').split('#')[0]
	const tag = '<base href="' + escapeAttr(href) + '">'
	if (/<head\b[^>]*>/i.test(html)) return html.replace(/<head\b([^>]*)>/i, '<head$1>' + tag)
	if (/<html\b[^>]*>/i.test(html)) return html.replace(/<html\b([^>]*)>/i, '<html$1><head>' + tag + '</head>')
	return '<!doctype html><html><head>' + tag + '</head><body>' + html + '</body></html>'
}

export function createHiddenFrame(html: string): SlideFrame {
	const iframe = document.createElement('iframe')
	iframe.setAttribute('aria-hidden', 'true')
	iframe.style.cssText = 'position:fixed;left:-100000px;top:0;width:1280px;height:720px;border:0;visibility:hidden;'
	iframe.width = '1280'
	iframe.height = '720'
	document.body.appendChild(iframe)
	return {
		iframe,
		load: () =>
			new Promise<void>((resolve, reject) => {
				iframe.onload = () => resolve()
				iframe.onerror = () => reject(new Error('Slide render frame could not be opened.'))
				iframe.srcdoc = withBaseHref(html)
			}),
		resize(width: number, height: number) {
			iframe.width = String(Math.round(width))
			iframe.height = String(Math.round(height))
			iframe.style.width = Math.round(width) + 'px'
			iframe.style.height = Math.round(height) + 'px'
		},
		close() {
			try {
				iframe.remove()
			} catch {
				/* already detached */
			}
		},
	}
}

export function wait(ms: number): Promise<void> {
	return new Promise((resolve) => setTimeout(resolve, ms))
}

// <iconify-icon> fetches its glyph from the icon CDN and renders it into a
// shadow-DOM <svg> asynchronously. The model extractor (addShadowSvg) can only
// capture an icon once that <svg> exists, so we must wait for every icon in the
// slide to finish rendering before reading the model - otherwise icons drop out
// unpredictably depending on per-slide network timing.
const iconRendered = (ic: IconElement) => !!(ic.shadowRoot && ic.shadowRoot.querySelector('svg'))

// Wait until the given icon elements have rendered, giving up early once no new
// icon has resolved for a while (invalid icon names never resolve, so we must
// not block on them for the full timeout).
export async function waitForIconEls(icons: IconElement[], timeoutMs?: number, stallMs?: number): Promise<void> {
	const deadline = Date.now() + (timeoutMs || 6000)
	let best = -1
	let lastProgress = Date.now()
	while (Date.now() < deadline) {
		const done = icons.filter(iconRendered).length
		if (done === icons.length) return
		if (done > best) {
			best = done
			lastProgress = Date.now()
		} else if (Date.now() - lastProgress > (stallMs || 1500)) return
		icons.forEach((ic) => {
			try {
				if (typeof ic.loadIcon === 'function') ic.loadIcon()
			} catch {
				/* ignore */
			}
		})
		await wait(100)
	}
}

async function waitForIcons(doc: Document, timeoutMs: number): Promise<void> {
	await waitForIconEls(Array.from(doc.querySelectorAll<IconElement>('iconify-icon')), timeoutMs, 1500)
}

export async function settleFrame(frame: SlideFrame): Promise<void> {
	const doc = frame.iframe.contentDocument
	if (!doc) throw new Error('Slide render frame was closed before it could be read.')
	await Promise.all(
		Array.from(doc.querySelectorAll<HTMLLinkElement>('link[rel~="stylesheet"]')).map((link) => {
			if (link.sheet) return Promise.resolve()
			return new Promise<void>((resolve) => {
				link.addEventListener('load', () => resolve(), { once: true })
				link.addEventListener('error', () => resolve(), { once: true })
				setTimeout(resolve, 3000)
			})
		})
	)
	try {
		// `doc.fonts` is what may be absent; `ready` is a promise, and testing a
		// promise for truthiness answers yes whatever it resolves to. `Document.fonts`
		// is declared as always present, so the cast is what lets the test be written
		// at all — the CSS Font Loading API is optional and older engines ship without
		// it.
		const fonts = doc.fonts as FontFaceSet | undefined
		if (fonts) await Promise.race([fonts.ready, wait(3000)])
	} catch {
		/* fonts API unavailable */
	}
	await waitForIcons(doc, 5000)
	await new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve())))
	await wait(120)
}
