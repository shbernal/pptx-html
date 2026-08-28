/**
 * Iconify resolution: fetch icon SVGs (or render via the Iconify custom element)
 * and inline them into slide HTML before the render frames are built, so per-slide
 * frames never depend on custom-element paint timing.
 *
 * The network fetch sits behind an injectable resolver.
 * `defaultResolveIcon(name) => Promise<SVGString|null>` is the default (an
 * `api.iconify.design` fetch); the consumer owns network/caching policy and tests
 * inject a static, offline resolver. The resolver only returns SVG *markup* — DOM
 * normalisation (1em box, currentColor) stays here in `protoFromSvgString`, shared
 * with the custom-element fallback.
 */

import { wait, waitForIconEls } from './frame'

/** Resolve an iconify icon name (e.g. `mdi:home`) to an SVG string, or null. */
export type IconResolver = (name: string) => Promise<string | null>

/** Normalised `<svg>` prototypes, keyed by icon name. */
type IconProtos = Record<string, SVGElement>

/** An `<iconify-icon>` element, narrowed to the shadow root the fallback reads. */
interface IconElement extends Element {
	shadowRoot: ShadowRoot | null
	loadIcon?: () => void
}

// Replace every <iconify-icon> in an HTML string with an inline <svg> taken from
// `svgByName`, carrying over the original element's class/style so it keeps its
// size (1em) and colour (currentColor). Unknown/unresolved icons are left as-is.
function inlineIconsInHtml(html: string, svgByName: IconProtos): string {
	const doc = new DOMParser().parseFromString('<div id="__icnroot">' + html + '</div>', 'text/html')
	const root = doc.getElementById('__icnroot')
	if (!root) return html
	Array.from(root.querySelectorAll('iconify-icon')).forEach((ic) => {
		const proto = svgByName[ic.getAttribute('icon') ?? '']
		if (!proto) return
		const svg = proto.cloneNode(true) as SVGElement
		const style = ic.getAttribute('style')
		const cls = ic.getAttribute('class')
		if (style) svg.setAttribute('style', (svg.getAttribute('style') || '') + ';' + style)
		if (cls) svg.setAttribute('class', cls)
		ic.replaceWith(svg)
	})
	return root.innerHTML
}

function findIconifyScriptSrc(headHTML: string): string {
	const doc = new DOMParser().parseFromString('<head>' + (headHTML || '') + '</head>', 'text/html')
	const script = Array.from(doc.querySelectorAll('script[src]')).find((el) =>
		/iconify/i.test(el.getAttribute('src') || '')
	)
	return script?.getAttribute('src') || ''
}

function loadMainScript(src: string): Promise<void> {
	const absolute = new URL(src, document.baseURI || window.location.href).href
	const existing = Array.from(document.scripts).find(
		(script) => script.src === absolute || script.getAttribute('src') === src
	)
	if (existing) return Promise.resolve()
	return new Promise((resolve, reject) => {
		const script = document.createElement('script')
		script.src = absolute
		script.async = true
		script.onload = () => resolve()
		script.onerror = () => reject(new Error('Iconify script could not be loaded for PPTX export.'))
		document.head.appendChild(script)
	})
}

async function ensureMainIconify(headHTML: string): Promise<boolean> {
	// `Window.customElements` is declared as always present and is not: the
	// registry is missing outside a secure context and in older engines, which is
	// the whole reason this function tests for it. Named once, at the type it
	// really has, rather than guarded three times against a type that says the
	// guard is pointless.
	const registry = window.customElements as CustomElementRegistry | undefined
	try {
		if (registry?.get('iconify-icon')) return true
		const src = findIconifyScriptSrc(headHTML)
		if (!src || !registry) return false
		await Promise.race([loadMainScript(src), wait(8000)])
		await Promise.race([registry.whenDefined('iconify-icon'), wait(8000)])
		return !!registry.get('iconify-icon')
	} catch {
		return false
	}
}

// Default icon resolver: fetch the raw SVG markup for an iconify icon
// name (`prefix:icon`) from api.iconify.design. Returns the SVG string, or null on
// any failure. Consumers/tests inject an alternative via `opts.resolveIcon`.
export async function defaultResolveIcon(name: string): Promise<string | null> {
	const idx = String(name || '').indexOf(':')
	if (idx <= 0) return null
	const prefix = name.slice(0, idx)
	const icon = name.slice(idx + 1)
	if (!prefix || !icon) return null
	try {
		const url =
			'https://api.iconify.design/' +
			encodeURIComponent(prefix) +
			'/' +
			encodeURIComponent(icon) +
			'.svg?height=1em&width=1em'
		const response = await fetch(url, { mode: 'cors' })
		if (!response.ok) return null
		return await response.text()
	} catch {
		return null
	}
}

// Normalise an inline <svg> element to a 1em box that inherits text colour, so it
// flows like the original <iconify-icon>. Mutates and returns the passed node.
function normalizeIconSvg(svg: SVGElement): SVGElement {
	if (!svg.getAttribute('xmlns')) svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg')
	svg.setAttribute('width', '1em')
	svg.setAttribute('height', '1em')
	if (!svg.getAttribute('fill')) svg.setAttribute('fill', 'currentColor')
	svg.style.cssText = 'vertical-align:middle;flex:0 0 auto'
	return svg
}

// Parse SVG markup (from a resolver) into a normalised, inline-ready <svg> proto,
// or null if it isn't valid SVG. Browser-only (DOMParser).
function protoFromSvgString(svgText: string | null): SVGElement | null {
	if (!svgText) return null
	try {
		const doc = new DOMParser().parseFromString(svgText, 'image/svg+xml')
		// No `!root`: an XML parse always produces a root element, a failed one
		// included — which is the `parsererror` the rest of this line looks for.
		const root = doc.documentElement
		if (root.tagName.toLowerCase() !== 'svg' || root.querySelector('parsererror')) return null
		return normalizeIconSvg(root.cloneNode(true) as SVGElement)
	} catch {
		return null
	}
}

// Resolve every icon in the deck once and inline SVGs before slide frames are
// created. The injected `resolveIcon` (default: `defaultResolveIcon`) avoids
// depending on Iconify custom-element paint timing; the custom element path
// remains as a fallback for anything the resolver could not return.
// Returns icon-inlined slide HTML, or null to fall back to the original slides.
export async function inlineDeckIcons(
	headHTML: string,
	slides: string[],
	resolveIcon?: IconResolver
): Promise<string[] | null> {
	const resolve = typeof resolveIcon === 'function' ? resolveIcon : defaultResolveIcon
	try {
		if (!slides.some((s) => /<iconify-icon/i.test(s))) return null
		const names = new Set<string>()
		slides.forEach((s) => {
			for (const [, name] of s.matchAll(/<iconify-icon\b[^>]*?\sicon=["']([^"']+)["']/gi)) {
				if (name) names.add(name)
			}
		})
		if (!names.size) return null
		const svgByName: IconProtos = {}
		await Promise.all(
			Array.from(names).map(async (name) => {
				const proto = protoFromSvgString(await resolve(name))
				if (proto) svgByName[name] = proto
			})
		)
		const missing = Array.from(names).filter((name) => !svgByName[name])
		if (!missing.length) return slides.map((s) => inlineIconsInHtml(s, svgByName))
		if (!(await ensureMainIconify(headHTML))) {
			return Object.keys(svgByName).length ? slides.map((s) => inlineIconsInHtml(s, svgByName)) : null
		}
		const host = document.createElement('div')
		host.setAttribute('aria-hidden', 'true')
		host.style.cssText =
			'position:fixed;left:-99999px;top:0;width:24px;height:24px;overflow:visible;font-size:24px;line-height:1;color:#000;pointer-events:none;'
		const els: IconElement[] = []
		const byName: Record<string, IconElement> = {}
		missing.forEach((name) => {
			const ic = document.createElement('iconify-icon') as IconElement
			ic.setAttribute('icon', name)
			host.appendChild(ic)
			els.push(ic)
			byName[name] = ic
		})
		document.body.appendChild(host)
		try {
			await waitForIconEls(els, 12000, 12000)
			Object.entries(byName).forEach(([name, element]) => {
				const svg = element.shadowRoot?.querySelector('svg')
				if (!svg) return
				svgByName[name] = normalizeIconSvg(svg.cloneNode(true) as SVGElement)
			})
			if (!Object.keys(svgByName).length) return null
			return slides.map((s) => inlineIconsInHtml(s, svgByName))
		} finally {
			host.remove()
		}
	} catch {
		return null
	}
}
