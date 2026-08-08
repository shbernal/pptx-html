// @ts-nocheck
/**
 * Hidden-iframe lifecycle and "settle" waits (fonts, stylesheets, icons) for the
 * browser extract path. Host-window code, not the stringified extractor payload.
 */

export function escapeAttr(value) {
	return String(value || '')
		.replace(/&/g, '&amp;')
		.replace(/"/g, '&quot;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
}

export function withBaseHref(html) {
	if (/<base\b/i.test(html)) return html
	const href = (document.baseURI || window.location.href || '').split('#')[0]
	const tag = '<base href="' + escapeAttr(href) + '">'
	if (/<head\b[^>]*>/i.test(html)) return html.replace(/<head\b([^>]*)>/i, '<head$1>' + tag)
	if (/<html\b[^>]*>/i.test(html)) return html.replace(/<html\b([^>]*)>/i, '<html$1><head>' + tag + '</head>')
	return '<!doctype html><html><head>' + tag + '</head><body>' + html + '</body></html>'
}

export function createHiddenFrame(html) {
	const iframe = document.createElement('iframe')
	iframe.setAttribute('aria-hidden', 'true')
	iframe.style.cssText = 'position:fixed;left:-100000px;top:0;width:1280px;height:720px;border:0;visibility:hidden;'
	iframe.width = '1280'
	iframe.height = '720'
	document.body.appendChild(iframe)
	return {
		iframe,
		load: () => new Promise((resolve, reject) => {
			iframe.onload = () => resolve()
			iframe.onerror = () => reject(new Error('Slide render frame could not be opened.'))
			iframe.srcdoc = withBaseHref(html)
		}),
		resize(width, height) {
			iframe.width = String(Math.round(width))
			iframe.height = String(Math.round(height))
			iframe.style.width = Math.round(width) + 'px'
			iframe.style.height = Math.round(height) + 'px'
		},
		close() { try { iframe.remove() } catch (e) { /* already detached */ } }
	}
}

export function wait(ms) {
	return new Promise((resolve) => setTimeout(resolve, ms))
}

// <iconify-icon> fetches its glyph from the icon CDN and renders it into a
// shadow-DOM <svg> asynchronously. The model extractor (addShadowSvg) can only
// capture an icon once that <svg> exists, so we must wait for every icon in the
// slide to finish rendering before reading the model - otherwise icons drop out
// unpredictably depending on per-slide network timing.
const iconRendered = (ic) => ic.shadowRoot && ic.shadowRoot.querySelector('svg')

// Wait until the given icon elements have rendered, giving up early once no new
// icon has resolved for a while (invalid icon names never resolve, so we must
// not block on them for the full timeout).
export async function waitForIconEls(icons, timeoutMs, stallMs) {
	const deadline = Date.now() + (timeoutMs || 6000)
	let best = -1
	let lastProgress = Date.now()
	while (Date.now() < deadline) {
		const done = icons.filter(iconRendered).length
		if (done === icons.length) return
		if (done > best) { best = done; lastProgress = Date.now() }
		else if (Date.now() - lastProgress > (stallMs || 1500)) return
		icons.forEach((ic) => { try { if (typeof ic.loadIcon === 'function') ic.loadIcon() } catch (e) { /* ignore */ } })
		await wait(100)
	}
}

async function waitForIcons(doc, timeoutMs) {
	await waitForIconEls(Array.from(doc.querySelectorAll('iconify-icon')), timeoutMs, 1500)
}

export async function settleFrame(frame) {
	const doc = frame.iframe.contentDocument
	await Promise.all(Array.from(doc.querySelectorAll('link[rel~="stylesheet"]')).map((link) => {
		if (link.sheet) return Promise.resolve()
		return new Promise((resolve) => {
			link.addEventListener('load', resolve, { once: true })
			link.addEventListener('error', resolve, { once: true })
			setTimeout(resolve, 3000)
		})
	}))
	try {
		if (doc.fonts && doc.fonts.ready) await Promise.race([doc.fonts.ready, wait(3000)])
	} catch (e) { /* fonts API unavailable */ }
	await waitForIcons(doc, 5000)
	await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))
	await wait(120)
}
