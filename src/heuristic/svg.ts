/**
 * SVG → editable vector (custGeom) transform.
 *
 * Browser-only host-realm glue (uses `DOMParser`/the SVG DOM): it post-processes
 * an extracted {@link SlideModel}, replacing `image` items whose source is an
 * inline SVG data URL (icons, decorative vector graphics) with editable
 * {@link PathItem}s, using the pure `svg-path` parser for the geometry. Unlike the
 * core `extractor` (which is stringified into the slide iframe), this runs in the
 * host window where normal module imports work.
 *
 * **Conservative by design:** anything it can't faithfully represent — a
 * `transform`, a gradient/pattern paint, `<use>`/`<image>`/`<text>`, a missing
 * viewBox — makes it leave that item as the original raster image. This keeps the
 * PNG raster as a no-regression fallback, which is why `vectorizeSvg` is opt-in
 * and off by default.
 */

import type { Item, PathItem, Rect, FreeformPoint } from './model'
import type { SlideModel } from './model'
import {
	parsePathData,
	rectToPathData,
	circleToPathData,
	ellipseToPathData,
	lineToPathData,
	pointsToPathData,
	type GeomSeg,
} from './svg-path'

const SVG_DATA_URL_RE = /^data:image\/svg\+xml/i

// Elements we know how to vectorize. Anything else with painted geometry forces a
// raster fallback so we never silently drop content.
const DRAWABLE = new Set(['path', 'rect', 'circle', 'ellipse', 'line', 'polygon', 'polyline'])
// Elements whose presence means we can't faithfully vectorize → fall back.
const UNSUPPORTED = new Set(['use', 'image', 'text', 'foreignobject', 'switch'])

/** Decode an `data:image/svg+xml` URL (base64 or URL-encoded) to SVG markup. */
function decodeSvgDataUrl(src: string): string | null {
	const comma = src.indexOf(',')
	if (comma < 0) return null
	const meta = src.slice(0, comma)
	const body = src.slice(comma + 1)
	try {
		if (/;base64/i.test(meta)) {
			return decodeURIComponent(escape(atob(body)))
		}
		return decodeURIComponent(body)
	} catch {
		return null
	}
}

/** Resolve an inherited presentation attribute (`fill`/`stroke`) up the ancestry. */
function inheritedAttr(el: Element, name: string, stopAt: Element): string | null {
	let node: Element | null = el
	while (node) {
		const v = node.getAttribute(name)
		if (v != null && v !== '') return v
		const styleVal = readStyleProp(node.getAttribute('style'), name)
		if (styleVal) return styleVal
		if (node === stopAt) break
		node = node.parentElement
	}
	return null
}

function readStyleProp(style: string | null, prop: string): string | null {
	if (!style) return null
	const m = new RegExp('(?:^|;)\\s*' + prop + '\\s*:\\s*([^;]+)', 'i').exec(style)
	return m?.[1]?.trim() ?? null
}

const NAMED_COLORS: Record<string, string> = {
	black: '000000',
	white: 'FFFFFF',
	red: 'FF0000',
	green: '008000',
	blue: '0000FF',
	gray: '808080',
	grey: '808080',
	silver: 'C0C0C0',
	transparent: 'none',
}

/** Normalize an SVG paint value to an uppercase 6-hex string, `'none'`, or null. */
function normalizeColor(value: string | null): string | null {
	if (!value) return null
	const v = value.trim().toLowerCase()
	if (v === 'none' || v === 'transparent') return 'none'
	if (v === 'currentcolor') return null // should already be resolved by the extractor
	if (v.startsWith('url(')) return null // gradient/pattern — caller treats as unsupported
	// Each branch reads its capture off the match rather than indexing into it, so
	// the shorthand expansion and the channel mapping are total by construction
	// instead of by trusting the pattern to have matched what it says it did.
	const short = /^#([0-9a-f]{3})$/i.exec(v)?.[1]
	if (short !== undefined) return short.replace(/./g, (digit) => digit + digit).toUpperCase()
	const full = /^#([0-9a-f]{6})$/i.exec(v)?.[1]
	if (full !== undefined) return full.toUpperCase()
	const rgb = /^rgba?\(\s*([\d.]+)[\s,]+([\d.]+)[\s,]+([\d.]+)/i.exec(v)
	if (rgb) {
		const hex = (n: string) => Math.max(0, Math.min(255, Math.round(parseFloat(n)))).toString(16).padStart(2, '0')
		return rgb.slice(1, 4).map(hex).join('').toUpperCase()
	}
	if (NAMED_COLORS[v]) return NAMED_COLORS[v]
	return null
}

interface ViewBox {
	minX: number
	minY: number
	w: number
	h: number
}

/**
 * A whitespace/comma-separated list read as exactly four finite numbers, or null
 * when it is anything else. Returning a tuple rather than an array is the point:
 * the four values are then readable without a bounds check each, and "the list
 * was the right length" is stated once, where it is decided.
 */
function fourNumbers(text: string): [number, number, number, number] | null {
	const parts = text.trim().split(/[\s,]+/).map(Number)
	const [a, b, c, d] = parts
	if (parts.length !== 4) return null
	if (a === undefined || b === undefined || c === undefined || d === undefined) return null
	if (!Number.isFinite(a) || !Number.isFinite(b) || !Number.isFinite(c) || !Number.isFinite(d)) return null
	return [a, b, c, d]
}

function readViewBox(svg: SVGSVGElement): ViewBox | null {
	const vb = svg.getAttribute('viewBox')
	if (vb) {
		const n = fourNumbers(vb)
		if (n && n[2] > 0 && n[3] > 0) {
			return { minX: n[0], minY: n[1], w: n[2], h: n[3] }
		}
	}
	const w = parseFloat(svg.getAttribute('width') || '')
	const h = parseFloat(svg.getAttribute('height') || '')
	if (Number.isFinite(w) && Number.isFinite(h) && w > 0 && h > 0) {
		return { minX: 0, minY: 0, w, h }
	}
	return null
}

/** The element's own `d` (or a `d` synthesized from a basic shape), or null. */
function shapeToPathData(el: Element): string | null {
	const tag = el.tagName.toLowerCase()
	const a = (n: string) => el.getAttribute(n)
	switch (tag) {
		case 'path':
			return a('d')
		case 'rect':
			return rectToPathData(a('x'), a('y'), a('width'), a('height'), a('rx'), a('ry'))
		case 'circle':
			return circleToPathData(a('cx'), a('cy'), a('r'))
		case 'ellipse':
			return ellipseToPathData(a('cx'), a('cy'), a('rx'), a('ry'))
		case 'line':
			return lineToPathData(a('x1'), a('y1'), a('x2'), a('y2'))
		case 'polygon':
			return pointsToPathData(a('points'), true)
		case 'polyline':
			return pointsToPathData(a('points'), false)
		default:
			return null
	}
}

// Map a uniform-or-not viewBox→box transform onto a point, honoring the default
// `preserveAspectRatio` (xMidYMid meet) so the vector matches how the browser
// rasterized the same SVG into the image we're replacing.
function makeMapper(vb: ViewBox, box: Rect, preserve: string | null) {
	const none = (preserve || '').trim().toLowerCase().startsWith('none')
	const sx = box.w / vb.w
	const sy = box.h / vb.h
	if (none) {
		return (x: number, y: number): [number, number] => [(x - vb.minX) * sx, (y - vb.minY) * sy]
	}
	const s = Math.min(sx, sy)
	const offX = (box.w - vb.w * s) / 2
	const offY = (box.h - vb.h * s) / 2
	return (x: number, y: number): [number, number] => [(x - vb.minX) * s + offX, (y - vb.minY) * s + offY]
}

function segsToPoints(segs: GeomSeg[], map: (x: number, y: number) => [number, number]): FreeformPoint[] {
	const pts: FreeformPoint[] = []
	for (const seg of segs) {
		if (seg.cmd === 'move') {
			const [x, y] = map(seg.x, seg.y)
			pts.push({ x, y, moveTo: true })
		} else if (seg.cmd === 'line') {
			const [x, y] = map(seg.x, seg.y)
			pts.push({ x, y })
		} else if (seg.cmd === 'cubic') {
			const [x1, y1] = map(seg.x1, seg.y1)
			const [x2, y2] = map(seg.x2, seg.y2)
			const [x, y] = map(seg.x, seg.y)
			pts.push({ x, y, curve: { type: 'cubic', x1, y1, x2, y2 } })
		} else if (seg.cmd === 'quad') {
			const [x1, y1] = map(seg.x1, seg.y1)
			const [x, y] = map(seg.x, seg.y)
			pts.push({ x, y, curve: { type: 'quadratic', x1, y1 } })
		} else if (seg.cmd === 'close') {
			pts.push({ close: true })
		}
	}
	return pts
}

/**
 * Convert one SVG image item to path items, or null if it can't be faithfully
 * vectorized (caller then keeps the raster image). Each painted drawable element
 * becomes one PathItem sharing the image's slide-space box.
 */
function svgItemToPaths(item: Extract<Item, { type: 'image' }>): PathItem[] | null {
	const markup = decodeSvgDataUrl(item.src)
	if (!markup) return null
	let svg: SVGSVGElement | null
	try {
		const doc = new DOMParser().parseFromString(markup, 'image/svg+xml')
		if (doc.querySelector('parsererror')) return null
		const root = doc.documentElement
		svg = root && root.tagName.toLowerCase() === 'svg' ? (root as unknown as SVGSVGElement) : null
	} catch {
		return null
	}
	if (!svg) return null

	const vb = readViewBox(svg)
	if (!vb) return null
	const map = makeMapper(vb, item.position, svg.getAttribute('preserveAspectRatio'))

	// Walk the tree in document order. Bail on anything we can't represent so the
	// raster fallback covers it, rather than emit wrong or partial geometry.
	const drawables: Element[] = []
	let unsupported = false
	const visit = (el: Element) => {
		if (unsupported) return
		const tag = el.tagName.toLowerCase()
		if (tag === 'defs' || tag === 'clippath' || tag === 'mask' || tag === 'symbol') return // ignore non-rendered subtrees
		if (UNSUPPORTED.has(tag)) {
			unsupported = true
			return
		}
		if (tag === 'svg' && el !== (svg)) {
			unsupported = true // nested SVG
			return
		}
		const transform = el.getAttribute('transform')
		if (transform && transform.trim()) {
			unsupported = true
			return
		}
		if (DRAWABLE.has(tag)) drawables.push(el)
		for (let c = el.firstElementChild; c; c = c.nextElementSibling) visit(c)
	}
	for (let c = svg.firstElementChild; c; c = c.nextElementSibling) visit(c)
	if (unsupported || !drawables.length) return null

	const paths: PathItem[] = []
	for (const el of drawables) {
		const d = shapeToPathData(el)
		if (!d) continue
		const fillRaw = inheritedAttr(el, 'fill', svg)
		const strokeRaw = inheritedAttr(el, 'stroke', svg)
		// A url() paint (gradient/pattern) can't be represented — fall back wholesale.
		if ((fillRaw && fillRaw.trim().toLowerCase().startsWith('url(')) || (strokeRaw && strokeRaw.trim().toLowerCase().startsWith('url(')))
			return null
		const fill = normalizeColor(fillRaw)
		const stroke = normalizeColor(strokeRaw)
		const segs = parsePathData(d)
		if (!segs.length) continue
		const points = segsToPoints(segs, map)
		if (!points.length) continue

		const path: PathItem = { type: 'path', z: item.z, position: { ...item.position }, points }
		// SVG default fill is black; an explicit `fill="none"` leaves it unfilled.
		if (fillRaw == null) path.fill = '000000'
		else if (fill && fill !== 'none') path.fill = fill
		if (stroke && stroke !== 'none') {
			const widthUser = parseFloat(inheritedAttr(el, 'stroke-width', svg) || '1') || 1
			// Convert stroke width from user units to points via the uniform scale.
			const s = Math.min(item.position.w / vb.w, item.position.h / vb.h)
			path.line = { color: stroke, width: Math.max(0.25, widthUser * s * 72) }
		}
		// A path with neither fill nor stroke paints nothing — skip it.
		if (!path.fill && !path.line) continue
		paths.push(path)
	}
	return paths.length ? paths : null
}

/**
 * In-place: replace every vectorizable SVG `image` item in the model with editable
 * `path` items. Returns the number of source images vectorized. Items that can't
 * be faithfully vectorized are left untouched (raster fallback). Called from the
 * deck orchestrator when `opts.vectorizeSvg` is set.
 */
export function vectorizeSvgImages(model: SlideModel): number {
	let converted = 0
	const out: Item[] = []
	for (const item of model.items) {
		if (item.type === 'image' && typeof item.src === 'string' && SVG_DATA_URL_RE.test(item.src)) {
			const paths = svgItemToPaths(item)
			if (paths) {
				out.push(...paths)
				converted++
				continue
			}
		}
		out.push(item)
	}
	model.items = out
	return converted
}
