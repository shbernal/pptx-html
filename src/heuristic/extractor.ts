// @ts-nocheck
/**
 * The browser model extractor — DOM → the heuristic lane's model.
 *
 * IMPORTANT: this function is **stringified and `eval`'d inside the slide
 * iframe** (see `./read.ts`). It must therefore stay entirely self-contained: it
 * may only reference its own nested helpers and DOM globals, never module-scope
 * imports. That is also why PX_PER_IN / PT_PER_PX are redefined here instead of
 * imported from `../constants`.
 *
 * **The one `@ts-nocheck` in the package, and it is quarantined rather than
 * pending.** It runs in a realm with no modules and returns across an `eval`
 * boundary as `unknown`, so nothing here can be checked where it executes.
 * `./read.ts` validates the whole result against `./model.ts` before any of it
 * reaches the writer — which is what keeps the suppression confined to the one
 * function that earns it instead of spreading downstream. Retyping this file in
 * place would not change what is checked at the point that matters; adding a
 * second unchecked consumer of its output would.
 *
 * Kept byte-for-byte faithful to the original browser engine so output stays
 * identical.
 */
export async function browserExtractor(config) {
	const PX_PER_IN = 96
	const PT_PER_PX = 0.75
	const fonts = config.fonts || {}
	const model = { background: { type: 'color', value: 'FFFFFF' }, items: [], notes: '' }
	const processed = new Set()
	const textRoots = new Set(['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'BLOCKQUOTE', 'FIGCAPTION'])
	const containerRoots = new Set(['DIV', 'SECTION', 'ARTICLE', 'HEADER', 'FOOTER', 'MAIN', 'ASIDE', 'NAV', 'FIGURE', 'SPAN', 'A', 'LABEL', 'MARK'])

	document.querySelectorAll('*').forEach((el) => { el.style.boxSizing = 'border-box' })
	void document.body.offsetHeight

	const root = document.querySelector('.slide') || document.body
	const rootRect = root.getBoundingClientRect()
	const scale = config.slideWidthIn / ((rootRect.width || root.scrollWidth || 1280) / PX_PER_IN)
	const originX = rootRect.left || 0
	const originY = rootRect.top || 0

	const rel = (rect) => ({
		x: ((rect.left - originX) / PX_PER_IN) * scale,
		y: ((rect.top - originY) / PX_PER_IN) * scale,
		w: (rect.width / PX_PER_IN) * scale,
		h: (rect.height / PX_PER_IN) * scale
	})
	const pt = (px) => (parseFloat(px) || 0) * PT_PER_PX * scale
	const text = (node) => {
		const cs = getComputedStyle(node)
		return applyTextTransform(normalizeText(node.textContent || '', cs.whiteSpace), cs.textTransform)
	}
	const zIndex = (el) => {
		let out = 0
		for (let node = el; node && node !== document.body; node = node.parentElement) {
			const n = parseInt(getComputedStyle(node).zIndex, 10)
			if (Number.isFinite(n)) out = Math.max(out, n)
		}
		return out
	}

	function normalizeText(value, whiteSpace) {
		const s = String(value || '').replace(/\u00a0/g, ' ')
		if (/pre/.test(whiteSpace || '')) return s.replace(/[ \t]+\n/g, '\n').trim()
		return s.replace(/\s+/g, ' ').trim()
	}

	function colorToHex(value, fallback) {
		const v = String(value || '').trim()
		if (!v || v === 'transparent' || v === 'rgba(0, 0, 0, 0)') return fallback || 'FFFFFF'
		if (v[0] === '#') {
			const hex = v.slice(1)
			if (hex.length === 3) return hex.split('').map((c) => c + c).join('').toUpperCase()
			return hex.slice(0, 6).toUpperCase()
		}
		const m = v.match(/rgba?\(([^)]+)\)/i)
		if (!m) return fallback || '000000'
		return m[1].split(',').slice(0, 3).map((part) => {
			const n = Math.max(0, Math.min(255, parseFloat(part) || 0))
			return Math.round(n).toString(16).padStart(2, '0')
		}).join('').toUpperCase()
	}

	function alphaTransparency(color, opacity) {
		let alpha = 1
		const m = String(color || '').match(/rgba?\(([^)]+)\)/i)
		if (m) {
			const parts = m[1].split(',').map((p) => p.trim())
			if (parts.length >= 4) alpha = parseFloat(parts[3])
		}
		const op = parseFloat(opacity)
		if (Number.isFinite(op)) alpha *= op
		if (alpha >= 0.995) return null
		return Math.max(0, Math.min(100, Math.round((1 - alpha) * 100)))
	}

	function isBoldWeight(weight) {
		const value = String(weight || '').trim().toLowerCase()
		if (value === 'bold' || value === 'bolder') return true
		const numeric = parseInt(value, 10)
		return Number.isFinite(numeric) && numeric >= 700
	}

	function fontFamilies(family) {
		const out = []
		const value = String(family || '')
		let current = ''
		let quote = ''
		for (let i = 0; i < value.length; i++) {
			const ch = value[i]
			if ((ch === '"' || ch === '\'') && !quote) {
				quote = ch
				continue
			}
			if (quote && ch === quote) {
				quote = ''
				continue
			}
			if (ch === ',' && !quote) {
				if (current.trim()) out.push(current.trim())
				current = ''
				continue
			}
			current += ch
		}
		if (current.trim()) out.push(current.trim())
		return out.map((name) => name.replace(/^["']|["']$/g, '').trim()).filter(Boolean)
	}

	function mapFont(family, sample) {
		const names = fontFamilies(family)
		const lowerNames = names.map((name) => name.toLowerCase())
		if (/[\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af]/.test(sample || '')) return fonts.cjk || 'Microsoft YaHei'
		if (lowerNames.some((name) => /material icons|fontawesome|bootstrap-icons|symbol/.test(name))) return fonts.symbol || 'Segoe UI Symbol'
		const generic = /^(serif|sans-serif|monospace|cursive|fantasy|system-ui|ui-serif|ui-sans-serif|ui-monospace|emoji|math|fangsong)$/
		const explicit = names.find((name) => !generic.test(name.toLowerCase()))
		if (explicit) return explicit
		const lower = lowerNames[0] || ''
		if (/bebas|oswald|anton|condensed|narrow|impact/.test(lower)) return fonts.emphasis || 'Liberation Sans Narrow'
		if (/serif|georgia|times|garamond|playfair/.test(lower)) return 'Tinos'
		if (/mono|code|courier/.test(lower)) return 'Courier New'
		return fonts.latin || 'Carlito'
	}

	function firstGradientLayer(value) {
		const s = String(value || '')
		const start = s.search(/(?:repeating-)?(?:linear|radial)-gradient\(/i)
		if (start < 0) return s
		const open = s.indexOf('(', start)
		if (open < 0) return s.slice(start)
		let depth = 0
		for (let i = open; i < s.length; i++) {
			if (s[i] === '(') depth++
			else if (s[i] === ')') {
				depth--
				if (depth === 0) return s.slice(start, i + 1)
			}
		}
		return s.slice(start)
	}

	function splitGradientArgs(value) {
		const inner = String(value || '').replace(/^[^(]*\(/, '').replace(/\)\s*$/, '')
		const out = []
		let depth = 0
		let start = 0
		for (let i = 0; i < inner.length; i++) {
			const ch = inner[i]
			if (ch === '(') depth++
			else if (ch === ')') depth = Math.max(0, depth - 1)
			else if (ch === ',' && depth === 0) {
				out.push(inner.slice(start, i).trim())
				start = i + 1
			}
		}
		out.push(inner.slice(start).trim())
		return out.filter(Boolean)
	}

	function extractStopColor(stop) {
		const m = String(stop || '').match(/rgba?\([^)]+\)|hsla?\([^)]+\)|#[0-9a-fA-F]{3,8}\b|\btransparent\b/i)
		return m ? m[0] : null
	}

	function parseGradientStops(args) {
		const stopArgs = args.filter((arg) => extractStopColor(arg))
		const stops = stopArgs.map((arg, index) => {
			const color = extractStopColor(arg)
			const tail = arg.slice(arg.indexOf(color) + color.length)
			const pct = tail.match(/(-?[\d.]+)%/)
			return {
				color: /^transparent$/i.test(color) ? 'rgba(0,0,0,0)' : color,
				pos: pct ? Math.max(0, Math.min(1, parseFloat(pct[1]) / 100)) : null,
				index
			}
		})
		if (!stops.length) return []
		if (stops.length === 1) stops[0].pos = 0
		stops.forEach((stop, index) => {
			if (stop.pos == null) stop.pos = stops.length === 1 ? 0 : index / (stops.length - 1)
		})
		return stops.toSorted((a, b) => a.pos - b.pos || a.index - b.index)
	}

	function gradientLine(width, height, args) {
		const header = args.find((arg) => !extractStopColor(arg)) || ''
		const angleMatch = header.match(/(-?[\d.]+)deg/i)
		let angle = angleMatch ? parseFloat(angleMatch[1]) : 135
		if (/to\s+right/i.test(header)) angle = 90
		else if (/to\s+left/i.test(header)) angle = 270
		else if (/to\s+bottom/i.test(header)) angle = 180
		else if (/to\s+top/i.test(header)) angle = 0
		const rad = angle * Math.PI / 180
		const dx = Math.sin(rad)
		const dy = -Math.cos(rad)
		const len = Math.abs(width * dx) + Math.abs(height * dy)
		const cx = width / 2
		const cy = height / 2
		return {
			x0: cx - dx * len / 2,
			y0: cy - dy * len / 2,
			x1: cx + dx * len / 2,
			y1: cy + dy * len / 2
		}
	}

	function gradientCenter(width, height, args) {
		const header = args.find((arg) => !extractStopColor(arg)) || ''
		const at = (header.match(/\bat\s+(.+)$/i) || [])[1] || ''
		let x = 0.5
		let y = 0.5
		const percents = Array.from(at.matchAll(/(-?[\d.]+)%/g)).map((m) => parseFloat(m[1]) / 100)
		if (percents.length >= 1) x = percents[0]
		if (percents.length >= 2) y = percents[1]
		if (/\bleft\b/i.test(at)) x = 0
		else if (/\bright\b/i.test(at)) x = 1
		if (/\btop\b/i.test(at)) y = 0
		else if (/\bbottom\b/i.test(at)) y = 1
		return {
			x: Math.max(0, Math.min(1, x)) * width,
			y: Math.max(0, Math.min(1, y)) * height
		}
	}

	function renderGradient(css, width, height, blurPx) {
		const gradientCss = firstGradientLayer(css)
		const args = splitGradientArgs(gradientCss)
		const stops = parseGradientStops(args)
		if (!stops.length) return null
		const canvas = document.createElement('canvas')
		canvas.width = Math.max(1, Math.round(width))
		canvas.height = Math.max(1, Math.round(height))
		const ctx = canvas.getContext('2d')
		if (blurPx) {
			try { ctx.filter = 'blur(' + blurPx + 'px)' } catch (e) {}
		}
		let g
		if (/radial-gradient/i.test(gradientCss)) {
			const c = gradientCenter(canvas.width, canvas.height, args)
			const radius = Math.max(
				Math.hypot(c.x, c.y),
				Math.hypot(canvas.width - c.x, c.y),
				Math.hypot(c.x, canvas.height - c.y),
				Math.hypot(canvas.width - c.x, canvas.height - c.y)
			)
			g = ctx.createRadialGradient(c.x, c.y, 0, c.x, c.y, radius)
		} else {
			const line = gradientLine(canvas.width, canvas.height, args)
			g = ctx.createLinearGradient(line.x0, line.y0, line.x1, line.y1)
		}
		stops.forEach((stop) => g.addColorStop(stop.pos, stop.color))
		ctx.fillStyle = g
		ctx.fillRect(0, 0, canvas.width, canvas.height)
		return canvas.toDataURL('image/png')
	}

	function renderBlurredSolid(computed, rect, blurPx) {
		const pad = Math.ceil(blurPx * 2.5)
		const canvas = document.createElement('canvas')
		canvas.width = Math.max(1, Math.round(rect.width + pad * 2))
		canvas.height = Math.max(1, Math.round(rect.height + pad * 2))
		const ctx = canvas.getContext('2d')
		try { ctx.filter = 'blur(' + blurPx + 'px)' } catch (e) {}
		ctx.fillStyle = computed.backgroundColor
		const radius = parseFloat(computed.borderRadius) || 0
		roundRect(ctx, pad, pad, rect.width, rect.height, radius)
		ctx.fill()
		return { data: canvas.toDataURL('image/png'), pad }
	}

	function roundRect(ctx, x, y, w, h, r) {
		const rr = Math.max(0, Math.min(r || 0, w / 2, h / 2))
		ctx.beginPath()
		ctx.moveTo(x + rr, y)
		ctx.lineTo(x + w - rr, y)
		ctx.quadraticCurveTo(x + w, y, x + w, y + rr)
		ctx.lineTo(x + w, y + h - rr)
		ctx.quadraticCurveTo(x + w, y + h, x + w - rr, y + h)
		ctx.lineTo(x + rr, y + h)
		ctx.quadraticCurveTo(x, y + h, x, y + h - rr)
		ctx.lineTo(x, y + rr)
		ctx.quadraticCurveTo(x, y, x + rr, y)
		ctx.closePath()
	}

	function splitCssList(value) {
		const out = []
		let depth = 0
		let start = 0
		const s = String(value || '')
		for (let i = 0; i < s.length; i++) {
			const ch = s[i]
			if (ch === '(') depth++
			else if (ch === ')') depth = Math.max(0, depth - 1)
			else if (ch === ',' && depth === 0) {
				out.push(s.slice(start, i).trim())
				start = i + 1
			}
		}
		out.push(s.slice(start).trim())
		return out.filter(Boolean)
	}

	function firstBoxShadow(value) {
		const shadow = splitCssList(value).find((part) => part && part !== 'none' && !/\binset\b/i.test(part))
		if (!shadow) return null
		const color = (shadow.match(/rgba?\([^)]+\)|hsla?\([^)]+\)|#[0-9a-fA-F]{3,8}\b|\b[a-z]+\b/i) || [])[0]
		if (!color || /^none$/i.test(color) || /^transparent$/i.test(color)) return null
		const lengths = shadow.replace(color, '').match(/-?[\d.]+px|-?[\d.]+/g) || []
		if (lengths.length < 2) return null
		return {
			color,
			offsetX: parseFloat(lengths[0]) || 0,
			offsetY: parseFloat(lengths[1]) || 0,
			blur: parseFloat(lengths[2]) || 0,
			spread: parseFloat(lengths[3]) || 0
		}
	}

	function renderBoxShadow(computed, rect) {
		const shadow = firstBoxShadow(computed.boxShadow)
		if (!shadow || (!shadow.offsetX && !shadow.offsetY && !shadow.blur && !shadow.spread)) return null
		const radius = parseFloat(computed.borderRadius) || 0
		const pad = Math.ceil(Math.abs(shadow.blur) * 2.2 + Math.abs(shadow.spread) + Math.max(Math.abs(shadow.offsetX), Math.abs(shadow.offsetY), 2))
		const canvas = document.createElement('canvas')
		canvas.width = Math.max(1, Math.round(rect.width + pad * 2))
		canvas.height = Math.max(1, Math.round(rect.height + pad * 2))
		const ctx = canvas.getContext('2d')
		ctx.shadowColor = shadow.color
		ctx.shadowBlur = Math.max(0, shadow.blur)
		ctx.shadowOffsetX = shadow.offsetX
		ctx.shadowOffsetY = shadow.offsetY
		ctx.fillStyle = shadow.color
		const spread = shadow.spread || 0
		roundRect(ctx, pad - spread, pad - spread, rect.width + spread * 2, rect.height + spread * 2, radius)
		ctx.fill()
		return { data: canvas.toDataURL('image/png'), pad }
	}

	function extractBackground() {
		const cs = getComputedStyle(root)
		if (cs.backgroundImage && cs.backgroundImage !== 'none' && /gradient/i.test(cs.backgroundImage)) {
			const image = renderGradient(cs.backgroundImage, rootRect.width || 1280, rootRect.height || 720, 0)
			if (image) {
				model.background = { type: 'color', value: colorToHex(cs.backgroundColor, 'FFFFFF') }
				model.items.push({ type: 'image', src: image, objectFit: 'fill', z: -10000, kind: 'decor', position: rel(rootRect) })
				return
			}
		}
		const url = String(cs.backgroundImage || '').match(/url\(["']?([^"')]+)["']?\)/)
		if (url) model.background = { type: 'image', src: url[1], fallback: colorToHex(cs.backgroundColor, 'FFFFFF') }
		else model.background = { type: 'color', value: colorToHex(cs.backgroundColor, 'FFFFFF') }
	}

	function collectNotes() {
		const notes = Array.from(document.querySelectorAll('[data-notes]')).map((node) => text(node)).filter(Boolean)
		model.notes = notes.join('\n')
		document.querySelectorAll('[data-notes]').forEach((node) => node.remove())
	}

	function addShadowSvg(el) {
		if (!el.shadowRoot || !el.shadowRoot.querySelector('svg')) return false
		const rect = el.getBoundingClientRect()
		if (rect.width <= 0 || rect.height <= 0) return mark(el)
		const svg = el.shadowRoot.querySelector('svg').cloneNode(true)
		const color = getComputedStyle(el).color || '#000'
		if (!svg.getAttribute('xmlns')) svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg')
		const svgNodes = [svg, ...Array.from(svg.querySelectorAll('*'))]
		svgNodes.forEach((node) => {
			['fill', 'stroke'].forEach((attr) => {
				if (node.getAttribute && node.getAttribute(attr) === 'currentColor') node.setAttribute(attr, color)
			})
		})
		if (!svg.getAttribute('fill')) svg.setAttribute('fill', color)
		svg.setAttribute('width', String(rect.width))
		svg.setAttribute('height', String(rect.height))
		model.items.push({ type: 'image', src: svgToDataUrl(svg), z: zIndex(el), position: rel(rect) })
		return mark(el)
	}

	function addInlineSvg(el) {
		if (!(el.tagName === 'svg' || el.tagName === 'SVG' || (el.namespaceURI === 'http://www.w3.org/2000/svg' && el.tagName.toLowerCase() === 'svg'))) return false
		const rect = el.getBoundingClientRect()
		if (rect.width > 0 && rect.height > 0) {
			const clone = el.cloneNode(true)
			const color = getComputedStyle(el).color || '#000'
			if (!clone.getAttribute('xmlns')) clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg')
			const cloneNodes = [clone, ...Array.from(clone.querySelectorAll('*'))]
			cloneNodes.forEach((node) => {
				['fill', 'stroke'].forEach((attr) => {
					if (node.getAttribute && node.getAttribute(attr) === 'currentColor') node.setAttribute(attr, color)
				})
			})
			clone.setAttribute('width', String(rect.width))
			clone.setAttribute('height', String(rect.height))
			model.items.push({ type: 'image', src: svgToDataUrl(clone), z: zIndex(el), position: rel(rect) })
		}
		return mark(el)
	}

	function svgToDataUrl(svg) {
		const text = new XMLSerializer().serializeToString(svg)
		return 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(text)))
	}

	async function normalizeImageSources() {
		const images = []
		if (model.background && model.background.src) images.push(model.background)
		model.items.forEach((item) => {
			if (item.type === 'image' && item.src) images.push(item)
		})
		await Promise.all(images.map(async (item) => {
			if (!item.src || /^data:/i.test(item.src) || /^blob:/i.test(item.src)) return
			try {
				const response = await fetch(item.src, { mode: 'cors' })
				if (!response.ok) return
				const blob = await response.blob()
				item.src = await new Promise((resolve, reject) => {
					const reader = new FileReader()
					reader.onload = () => resolve(reader.result)
					reader.onerror = () => reject(new Error('Image data could not be read.'))
					reader.readAsDataURL(blob)
				})
			} catch (_) {
				// Keep the original URL; PptxGen can still resolve many browser-visible paths.
			}
		}))
		await Promise.all(images.map((item) => cropCoverImage(item)))
	}

	function loadImageNode(src) {
		return new Promise((resolve, reject) => {
			const img = new Image()
			img.onload = () => resolve(img)
			img.onerror = () => reject(new Error('Image could not be prepared for PPTX export.'))
			img.src = src
		})
	}

	async function cropCoverImage(item) {
		if (!item || item.objectFit !== 'cover' || !item.position || !/^data:image\//i.test(item.src) || /^data:image\/svg/i.test(item.src)) return
		const targetAspect = item.position.w / item.position.h
		if (!Number.isFinite(targetAspect) || targetAspect <= 0) return
		try {
			const img = await loadImageNode(item.src)
			const iw = img.naturalWidth || img.width
			const ih = img.naturalHeight || img.height
			if (!iw || !ih) return
			const sourceAspect = iw / ih
			let sx = 0
			let sy = 0
			let sw = iw
			let sh = ih
			if (sourceAspect > targetAspect) {
				sw = ih * targetAspect
				sx = (iw - sw) / 2
			} else if (sourceAspect < targetAspect) {
				sh = iw / targetAspect
				sy = (ih - sh) / 2
			}
			const maxLongEdge = 1800
			const scale = Math.min(1, maxLongEdge / Math.max(sw, sh))
			const canvas = document.createElement('canvas')
			canvas.width = Math.max(1, Math.round(sw * scale))
			canvas.height = Math.max(1, Math.round(sh * scale))
			canvas.getContext('2d').drawImage(img, sx, sy, sw, sh, 0, 0, canvas.width, canvas.height)
			const mime = /^data:image\/jpe?g/i.test(item.src) ? 'image/jpeg' : 'image/png'
			item.src = mime === 'image/jpeg' ? canvas.toDataURL(mime, 0.92) : canvas.toDataURL(mime)
			item.objectFit = 'fill'
		} catch (_) {}
	}

	function addImage(el) {
		if (el.tagName !== 'IMG') return false
		const rect = el.getBoundingClientRect()
		if (rect.width > 0 && rect.height > 0) {
			const cs = getComputedStyle(el)
			const item = {
				type: 'image',
				src: el.src,
				objectFit: cs.objectFit === 'contain' ? 'contain' : (cs.objectFit === 'cover' ? 'cover' : null),
				objectPosition: cs.objectPosition,
				z: zIndex(el),
				position: rel(rect)
			}
			const t = alphaTransparency(null, cs.opacity)
			if (t !== null) item.transparency = t
			model.items.push(item)
		}
		return mark(el)
	}

	function addTable(el) {
		if (el.tagName !== 'TABLE') return false
		const rect = el.getBoundingClientRect()
		if (rect.width <= 0 || rect.height <= 0) return mark(el)
		const tableCs = getComputedStyle(el)
		const tableBg = visibleColor(tableCs.backgroundColor) ? tableCs.backgroundColor : null
		const rowEls = Array.from(el.querySelectorAll('tr')).filter((tr) => Array.from(tr.children).some((c) => c.tagName === 'TD' || c.tagName === 'TH'))
		const colW = measureTableColumns(rowEls, rect)
		const rowH = rowEls.map((tr) => rel(tr.getBoundingClientRect()).h).filter((h) => h > 0)
		const rows = rowEls.map((tr, rowIndex) => {
			const trCs = getComputedStyle(tr)
			const sectionCs = tr.parentElement ? getComputedStyle(tr.parentElement) : null
			const inheritedBg = firstVisibleColor(trCs.backgroundColor, sectionCs && sectionCs.backgroundColor, tableBg)
			return Array.from(tr.children).filter((c) => c.tagName === 'TD' || c.tagName === 'TH').map((cell, cellIndex) => {
				const cs = getComputedStyle(cell)
				const cellRect = cell.getBoundingClientRect()
				const textRect = textRangeRect(cell)
				const leftPadPx = textRect ? Math.max(parseFloat(cs.paddingLeft) || 0, textRect.left - cellRect.left) : (parseFloat(cs.paddingLeft) || 0)
				const opts = {
					bold: cell.tagName === 'TH' || isBoldWeight(cs.fontWeight),
					fontFace: mapFont(cs.fontFamily, tableCellText(cell)),
					fontSize: Math.max(7, pt(cs.fontSize)),
					color: colorToHex(cs.color, '000000'),
					align: /^(left|right|center|justify)$/.test(cs.textAlign) ? cs.textAlign : 'left',
					margin: [pt(cs.paddingTop), pt(cs.paddingRight), pt(cs.paddingBottom), pt(leftPadPx + 'px')]
				}
				const valign = mapVerticalAlign(cs.verticalAlign)
				if (valign) opts.valign = valign
				const bg = firstVisibleColor(cs.backgroundColor, inheritedBg)
				if (bg) {
					opts.fill = { color: colorToHex(bg, 'FFFFFF') }
					const t = alphaTransparency(bg, cs.opacity)
					if (t !== null) opts.fill.transparency = t
				}
				const border = tableCellBorder(cs, trCs, tableCs)
				if (border) opts.border = border
				const colspan = parseInt(cell.getAttribute('colspan') || '1', 10)
				const rowspan = parseInt(cell.getAttribute('rowspan') || '1', 10)
				if (colspan > 1) opts.colspan = colspan
				if (rowspan > 1) opts.rowspan = rowspan
				addTableCellOverlays(cell, rowIndex, cellIndex)
				return { text: tableCellBaseText(cell), options: opts }
			})
		}).filter((row) => row.length)
		if (rows.length) {
			model.items.push({
				type: 'table',
				rows,
				colW,
				rowH,
				z: zIndex(el),
				position: rel(rect),
				border: tableBorder(tableCs)
			})
		}
		return mark(el)
	}

	function visibleColor(value) {
		const v = String(value || '').trim()
		return !!(v && v !== 'transparent' && v !== 'rgba(0, 0, 0, 0)' && v !== 'rgba(0,0,0,0)')
	}

	function firstVisibleColor() {
		for (const value of arguments) {
			if (visibleColor(value)) return value
		}
		return null
	}

	function mapVerticalAlign(value) {
		const v = String(value || '').toLowerCase()
		if (v === 'middle') return 'middle'
		if (v === 'bottom') return 'bottom'
		if (v === 'top' || v === 'baseline' || v === 'text-top') return 'top'
		return null
	}

	function tableBorder(cs) {
		const side = borderSide(cs, 'Top')
		return side || hiddenBorderSide()
	}

	function tableCellBorder(cellCs, rowCs, tableCs) {
		const sides = ['Top', 'Right', 'Bottom', 'Left'].map((side) => {
			return borderSide(cellCs, side) || borderSide(rowCs, side) || borderSide(tableCs, side) || hiddenBorderSide()
		})
		if (sides.every((side) => side.transparency === 100 || side.pt === 0)) return null
		return sides
	}

	function hiddenBorderSide() {
		return { type: 'solid', color: 'FFFFFF', transparency: 100, pt: 0.25 }
	}

	function borderSide(cs, side) {
		if (!cs) return null
		const width = parseFloat(cs['border' + side + 'Width']) || 0
		const style = cs['border' + side + 'Style']
		const color = cs['border' + side + 'Color']
		if (width <= 0 || /none|hidden/i.test(style || '') || !visibleColor(color)) return null
		const transparency = alphaTransparency(color, null)
		if (isWeakDecorativeBorder(width, transparency)) return null
		const out = {
			type: 'solid',
			color: colorToHex(color, 'CCCCCC'),
			pt: Math.max(0.25, pt(width + 'px') || 0.5)
		}
		if (/dashed/i.test(style || '')) out.dash = 'dash'
		if (/dotted/i.test(style || '')) out.dash = 'dot'
		if (transparency !== null) out.transparency = transparency
		return out
	}

	function isWeakDecorativeBorder(widthPx, transparency) {
		return widthPx <= 2 && transparency !== null && transparency >= 45
	}

	function measureTableColumns(rowEls, tableRect) {
		const probe = rowEls.find((tr) => tr.children && tr.children.length)
		if (!probe || !tableRect.width) return null
		const widths = []
		Array.from(probe.children).filter((c) => c.tagName === 'TD' || c.tagName === 'TH').forEach((cell) => {
			const colspan = Math.max(1, parseInt(cell.getAttribute('colspan') || '1', 10))
			const w = rel(cell.getBoundingClientRect()).w / colspan
			for (let i = 0; i < colspan; i++) widths.push(w)
		})
		return widths.length ? widths : null
	}

	function isTableVisualOverlay(el) {
		if (!el || el.tagName === 'svg' || el.tagName === 'SVG' || el.tagName === 'ICONIFY-ICON') return false
		const tag = el.tagName.toLowerCase()
		if (!/^(span|mark|b|strong|em|i|small|label|div)$/i.test(tag)) return false
		const cs = getComputedStyle(el)
		const cls = String(el.className || '')
		const hasVisualClass = /\b(badge|chip|pill|tag|label|status|priority|domain|raci|crit|high|med|low|warn|risk)\b/i.test(cls)
		const hasBoxPaint = firstVisibleColor(cs.backgroundColor) || visibleBorder(cs)
		const inlineBox = /inline-flex|inline-block|flex|grid/i.test(cs.display || '')
		return !!text(el) && (hasVisualClass || hasBoxPaint || inlineBox)
	}

	function visibleBorder(cs) {
		return ['Top', 'Right', 'Bottom', 'Left'].some((side) => borderSide(cs, side))
	}

	function tableCellBaseText(cell) {
		const visualChildren = Array.from(cell.children).filter(isTableVisualOverlay)
		const directText = Array.from(cell.childNodes).some((node) => node.nodeType === Node.TEXT_NODE && normalizeText(node.textContent, getComputedStyle(cell).whiteSpace))
		if (visualChildren.length && !directText) {
			const allText = normalizeText(tableCellText(cell), getComputedStyle(cell).whiteSpace)
			const visualText = normalizeText(visualChildren.map((node) => tableCellText(node)).join(' '), getComputedStyle(cell).whiteSpace)
			if (allText && allText === visualText) return ''
		}
		return tableCellText(cell)
	}

	function tableCellText(el) {
		const pieces = []
		const pushText = (value, owner) => {
			const cs = getComputedStyle(owner || el)
			const s = applyTextTransform(normalizeText(value || '', cs.whiteSpace), cs.textTransform)
			if (s) pieces.push(s)
		}
		const walk = (node) => {
			if (node.nodeType === Node.TEXT_NODE) {
				pushText(node.textContent, node.parentElement || el)
				return
			}
			if (node.nodeType !== Node.ELEMENT_NODE) return
			const tag = node.tagName
			if (tag === 'BR') {
				pieces.push('\n')
				return
			}
			if (tag === 'SVG' || tag === 'ICONIFY-ICON') return
			const cs = getComputedStyle(node)
			const blockish = /^(block|flex|grid|list-item|table|table-row|table-cell)$/i.test(cs.display || '') && node !== el
			if (blockish && pieces.length && pieces[pieces.length - 1] !== '\n') pieces.push('\n')
			Array.from(node.childNodes).forEach(walk)
			if (blockish && pieces.length && pieces[pieces.length - 1] !== '\n') pieces.push('\n')
		}
		Array.from(el.childNodes).forEach(walk)
		return pieces.join(' ')
			.replace(/[ \t]*\n[ \t]*/g, '\n')
			.replace(/\n{3,}/g, '\n\n')
			.replace(/[ \t]{2,}/g, ' ')
			.trim()
	}

	function applyTextTransform(value, transform) {
		const s = String(value || '')
		const t = String(transform || '').toLowerCase()
		if (!s || t === 'none') return s
		if (t === 'uppercase') return s.toUpperCase()
		if (t === 'lowercase') return s.toLowerCase()
		if (t === 'capitalize') return s.replace(/\b([\p{L}\p{N}])/gu, (m) => m.toUpperCase())
		return s
	}

	function addTableCellOverlays(cell, rowIndex, cellIndex) {
		Array.from(cell.querySelectorAll('*')).filter(isTableVisualOverlay).forEach((node) => {
			const r = node.getBoundingClientRect()
			if (r.width <= 0 || r.height <= 0) return
			const cs = getComputedStyle(node)
			const layerZ = zIndex(node) + 0.05
			const bg = firstVisibleColor(cs.backgroundColor)
			const borders = ['Top', 'Right', 'Bottom', 'Left'].map((side) => borderSide(cs, side)).filter(Boolean)
			if (bg || borders.length) {
				const shape = {
					type: 'shape',
					z: layerZ,
					position: rel(r),
					fill: bg ? colorToHex(bg, 'FFFFFF') : null,
					transparency: bg ? alphaTransparency(bg, cs.opacity) : null,
					radius: parseFloat(cs.borderRadius) || 0,
					line: borders.length ? { color: borders[0].color, width: borders[0].pt || 0.5, transparency: borders[0].transparency } : null,
					label: 'table cell overlay ' + (rowIndex + 1) + '.' + (cellIndex + 1)
				}
				model.items.push(shape)
			}
			model.items.push({
				type: 'text',
				tag: 'span',
				text: text(node),
				z: layerZ,
				noWrap: true,
				position: rel(r),
				style: textStyle(node, cs, text(node))
			})
		})
	}

	function addRule(el) {
		if (el.tagName !== 'HR') return false
		const rect = el.getBoundingClientRect()
		if (rect.width > 0) {
			const cs = getComputedStyle(el)
			model.items.push({
				type: 'line',
				color: colorToHex(cs.borderTopColor || cs.color, '888888'),
				width: Math.max(0.5, pt(cs.borderTopWidth) || 1),
				dash: cs.borderTopStyle === 'dashed' ? 'dash' : (cs.borderTopStyle === 'dotted' ? 'dot' : 'solid'),
				z: zIndex(el),
				position: rel({ left: rect.left, top: rect.top + rect.height / 2, width: rect.width, height: 0 })
			})
		}
		return mark(el)
	}

	function addContainer(el) {
		if (!containerRoots.has(el.tagName) || el === root || el.closest('p,h1,h2,h3,h4,h5,h6,li,table')) return false
		const rect = el.getBoundingClientRect()
		if (rect.width <= 0 || rect.height <= 0) return false
		const cs = getComputedStyle(el)
		const bg = cs.backgroundColor && cs.backgroundColor !== 'rgba(0, 0, 0, 0)' && cs.backgroundColor !== 'transparent'
		const bgImage = cs.backgroundImage && cs.backgroundImage !== 'none'
		const borderSides = [
			{ css: 'Top', side: 'top' },
			{ css: 'Right', side: 'right' },
			{ css: 'Bottom', side: 'bottom' },
			{ css: 'Left', side: 'left' }
		].map((side) => {
			const border = borderSide(cs, side.css)
			if (!border) return null
			return {
				...border,
				side: side.side,
				width: parseFloat(cs['border' + side.css + 'Width']) || 0
			}
		}).filter(Boolean)
		const border = borderSides.length > 0
		const partialBorder = border && borderSides.length < 4
		// All of a container's layers share its stacking z (zIndex). Their relative
		// order - fill, then background image, then gradient, then borders/shape, then
		// the element's own text and descendants - is preserved by push order, which
		// sortedItems() uses as the tiebreaker. Earlier builds nudged each layer down
		// by a fixed fraction (layerZ - 0.18..0.22); that collapsed DOM depth, so an
		// ancestor's background (e.g. a hero gradient) could end up above a descendant
		// overlay or a full-bleed photo. Keeping every layer at layerZ and trusting
		// document order matches how the browser actually paints these slides.
		const layerZ = zIndex(el)
		const blur = (String(cs.filter || '').match(/blur\(\s*([\d.]+)px\s*\)/i) || [])[1]
		const shadow = renderBoxShadow(cs, rect)
		if (shadow) {
			model.items.push({
				type: 'image',
				src: shadow.data,
				objectFit: 'fill',
				z: layerZ,
				kind: 'decor',
				transparency: alphaTransparency(null, cs.opacity),
				position: rel({ left: rect.left - shadow.pad, top: rect.top - shadow.pad, width: rect.width + shadow.pad * 2, height: rect.height + shadow.pad * 2 })
			})
		}
		if (blur && bg && !text(el)) {
			const out = renderBlurredSolid(cs, rect, parseFloat(blur))
			model.items.push({
				type: 'image',
				src: out.data,
				objectFit: 'fill',
				z: layerZ,
				kind: 'decor',
				transparency: alphaTransparency(null, cs.opacity),
				position: rel({ left: rect.left - out.pad, top: rect.top - out.pad, width: rect.width + out.pad * 2, height: rect.height + out.pad * 2 })
			})
			return mark(el)
		}
		if (bg && bgImage) {
			model.items.push({
				type: 'shape',
				z: layerZ,
				position: rel(rect),
				fill: colorToHex(cs.backgroundColor, 'FFFFFF'),
				transparency: alphaTransparency(cs.backgroundColor, cs.opacity),
				radius: parseFloat(cs.borderRadius) || 0,
				line: null,
				label: text(el).slice(0, 32)
			})
		}
		if (bgImage) {
			const urls = Array.from(String(cs.backgroundImage).matchAll(/url\(["']?([^"')]+)["']?\)/g)).map((match) => match[1])
			urls.forEach((src) => {
				model.items.push({
					type: 'image',
					src,
					objectFit: String(cs.backgroundSize || '').includes('contain') ? 'contain' : 'cover',
					z: layerZ,
					position: rel(rect)
				})
			})
		}
		if (bgImage && /gradient/i.test(cs.backgroundImage)) {
			const image = renderGradient(cs.backgroundImage, rect.width, rect.height, blur ? parseFloat(blur) : 0)
			if (image) model.items.push({ type: 'image', src: image, objectFit: 'fill', z: layerZ, kind: 'decor', transparency: alphaTransparency(null, cs.opacity), position: rel(rect) })
		}
		if ((bg && !bgImage) || (border && !partialBorder)) {
			const item = {
				type: 'shape',
				z: layerZ,
				position: rel(rect),
				fill: bg && !bgImage ? colorToHex(cs.backgroundColor, 'FFFFFF') : null,
				transparency: bg && !bgImage ? alphaTransparency(cs.backgroundColor, cs.opacity) : null,
				radius: parseFloat(cs.borderRadius) || 0,
				line: border && !partialBorder ? {
					color: borderSides[0].color,
					width: borderSides[0].pt || Math.max(0.25, pt(borderSides[0].width + 'px') || 0.5),
					transparency: borderSides[0].transparency
				} : null,
				label: text(el).slice(0, 32)
			}
			model.items.push(item)
		}
		if (partialBorder) {
			borderSides.forEach((side) => {
				const thickness = Math.max(1, side.width)
				const stripe = {
					type: 'shape',
					z: layerZ + 0.01,
					position: null,
					fill: side.color,
					transparency: side.transparency,
					radius: 0,
					line: null,
					label: 'border-' + side.side
				}
				if (side.side === 'top') stripe.position = rel({ left: rect.left, top: rect.top, width: rect.width, height: thickness })
				if (side.side === 'right') stripe.position = rel({ left: rect.right - thickness, top: rect.top, width: thickness, height: rect.height })
				if (side.side === 'bottom') stripe.position = rel({ left: rect.left, top: rect.bottom - thickness, width: rect.width, height: thickness })
				if (side.side === 'left') stripe.position = rel({ left: rect.left, top: rect.top, width: thickness, height: rect.height })
				model.items.push(stripe)
			})
		}
		const directText = Array.from(el.childNodes).some((n) => n.nodeType === Node.TEXT_NODE && n.textContent.trim())
		if (directText) {
			el._qualityTreatAsText = true
			return false
		}
		if (Array.from(el.children).some((child) => text(child))) {
			processed.add(el)
			return false
		}
		// A text-free wrapper can still hold real visual content - most commonly a
		// background photo with its overlay masks (<div><img><div class="mask">).
		// mark(el) would flag every descendant as processed and silently drop them,
		// so leave the wrapper unmarked and let its children be emitted individually.
		if (el.querySelector('img, svg, picture, canvas, video, iconify-icon')) {
			processed.add(el)
			return false
		}
		if (hasVisualDescendant(el)) {
			processed.add(el)
			return false
		}
		return mark(el)
	}

	function hasVisualDescendant(el) {
		return Array.from(el.querySelectorAll('*')).some((node) => {
			if (node.closest('p,h1,h2,h3,h4,h5,h6,li,table')) return false
			const r = node.getBoundingClientRect()
			if (r.width <= 0 || r.height <= 0) return false
			const cs = getComputedStyle(node)
			const hasBg = visibleColor(cs.backgroundColor) || (cs.backgroundImage && cs.backgroundImage !== 'none')
			const hasBorder = ['Top', 'Right', 'Bottom', 'Left'].some((side) => borderSide(cs, side))
			return hasBg || hasBorder
		})
	}

	function addList(el) {
		if (el.tagName !== 'UL' && el.tagName !== 'OL') return false
		if (el.parentElement && el.parentElement.closest('ul,ol')) return mark(el)
		const rect = el.getBoundingClientRect()
		if (rect.width <= 0 || rect.height <= 0) return mark(el)
		const ordered = el.tagName === 'OL'
		const items = []
		Array.from(el.children).filter((li) => li.tagName === 'LI').forEach((li, index, all) => {
			const runs = inlineRuns(li, { breakLine: index !== all.length - 1 })
			if (runs.length) {
				runs[0].text = runs[0].text.replace(/^[\u2022\-\*\u25aa\u25b8\u25cb\u25cf]\s*/, '').replace(/^\d+[\.\)]\s+/, '')
				runs[0].options.bullet = ordered ? { type: 'number' } : {}
				items.push(...runs)
			}
		})
		const cs = getComputedStyle(el)
		model.items.push({
			type: 'list',
			text: items,
			z: zIndex(el),
			position: rel(rect),
			style: textStyle(el, cs, text(el))
		})
		return mark(el)
	}

	function addText(el) {
		if (!textRoots.has(el.tagName) && !el._qualityTreatAsText) {
			const standalone = ['SPAN', 'A', 'LABEL', 'MARK', 'STRONG', 'B', 'EM', 'I', 'CODE'].includes(el.tagName) && !el.closest('p,h1,h2,h3,h4,h5,h6,li')
			if (!standalone) return false
		}
		if (el.closest('table')) return false
		const content = text(el)
		const rect = el.getBoundingClientRect()
		if (!content || rect.width <= 0 || rect.height <= 0) return false
		const cs = getComputedStyle(el)
		const fittedRect = el.querySelector('svg') ? (textRangeRect(el) || rect) : rect
		const hasRich = !!el.querySelector('strong,b,em,i,u,a,span,mark,code,small,big,br,sub,sup')
		model.items.push({
			type: 'text',
			tag: el._qualityTreatAsText ? 'p' : el.tagName.toLowerCase(),
			text: hasRich ? inlineRuns(el, {}) : content,
			z: zIndex(el),
			noWrap: cs.whiteSpace === 'nowrap' || cs.whiteSpace === 'pre' || (textLineCount(el) <= 1 && content.length <= 72),
			position: rel(fittedRect),
			style: textStyle(el, cs, content)
		})
		return mark(el)
	}

	function textRangeRect(el) {
		const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, {
			acceptNode(node) {
				if (!node.textContent || !node.textContent.trim()) return NodeFilter.FILTER_REJECT
				return node.parentElement && node.parentElement.closest('svg') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT
			}
		})
		let first = null
		let last = null
		while (walker.nextNode()) {
			if (!first) first = walker.currentNode
			last = walker.currentNode
		}
		if (!first || !last) return null
		const range = document.createRange()
		range.setStart(first, 0)
		range.setEnd(last, last.textContent.length)
		const rects = Array.from(range.getClientRects()).filter((r) => r.width > 0 && r.height > 0)
		range.detach()
		if (!rects.length) return null
		const left = Math.min(...rects.map((r) => r.left))
		const top = Math.min(...rects.map((r) => r.top))
		const right = Math.max(...rects.map((r) => r.right))
		const bottom = Math.max(...rects.map((r) => r.bottom))
		return { left, top, right, bottom, width: right - left, height: bottom - top }
	}

	function textLineCount(el) {
		const rect = textRangeRect(el)
		if (!rect) return 0
		const lineTops = []
		const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, {
			acceptNode(node) {
				if (!node.textContent || !node.textContent.trim()) return NodeFilter.FILTER_REJECT
				return node.parentElement && node.parentElement.closest('svg') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT
			}
		})
		while (walker.nextNode()) {
			const range = document.createRange()
			range.selectNodeContents(walker.currentNode)
			Array.from(range.getClientRects()).forEach((r) => {
				if (r.width <= 0 || r.height <= 0) return
				if (!lineTops.some((top) => Math.abs(top - r.top) < 2)) lineTops.push(r.top)
			})
			range.detach()
		}
		return lineTops.length || 1
	}

	function alignKeywordToPpt(value) {
		const v = String(value || '').toLowerCase()
		if (/^(center|middle)$/.test(v)) return 'middle'
		if (/^(flex-end|end|bottom)$/.test(v)) return 'bottom'
		if (/^(flex-start|start|top|normal|stretch)$/.test(v)) return 'top'
		return null
	}

	function textVerticalAlign(cs) {
		const display = String(cs.display || '').toLowerCase()
		if (/flex/.test(display)) {
			const dir = String(cs.flexDirection || '').toLowerCase()
			const mainAxis = /column/.test(dir) ? cs.justifyContent : cs.alignItems
			return alignKeywordToPpt(mainAxis) || 'top'
		}
		if (/grid/.test(display)) {
			return alignKeywordToPpt(cs.alignItems) || alignKeywordToPpt(cs.placeItems && cs.placeItems.split(/\s+/)[0]) || 'top'
		}
		return mapVerticalAlign(cs.verticalAlign) || 'top'
	}

	function shouldConstrainTextBox(cs) {
		const display = String(cs.display || '').toLowerCase()
		if (!/(flex|grid)/.test(display)) return false
		const centered = /center|middle/.test([cs.justifyContent, cs.alignItems, cs.placeItems].join(' '))
		return centered || visibleColor(cs.backgroundColor) || (cs.backgroundImage && cs.backgroundImage !== 'none') || visibleBorder(cs)
	}

	function textStyle(el, cs, sample) {
		const fillColor = textPaint(cs, el)
		const style = {
			fontFace: mapFont(cs.fontFamily, sample),
			fontSize: Math.max(7, pt(cs.fontSize)),
			color: fillColor,
			bold: isBoldWeight(cs.fontWeight),
			italic: cs.fontStyle === 'italic',
			underline: String(cs.textDecoration || '').includes('underline'),
			strike: String(cs.textDecoration || '').includes('line-through'),
			align: /^(left|right|center|justify)$/.test(cs.textAlign) ? cs.textAlign : 'left',
			valign: textVerticalAlign(cs),
			constrainTextBox: shouldConstrainTextBox(cs),
			margin: [pt(cs.paddingLeft), pt(cs.paddingRight), pt(cs.paddingBottom), pt(cs.paddingTop)],
			lineSpacing: cs.lineHeight && cs.lineHeight !== 'normal' ? pt(cs.lineHeight) : null,
			transparency: alphaTransparency(cs.color, cs.opacity)
		}
		if (/[\u2190-\u21ff\u2600-\u27bf]/.test(sample) && Array.from(sample.trim()).length <= 4) style.fontFace = fonts.symbol || 'Segoe UI Symbol'
		return style
	}

	function textPaint(cs, el) {
		if (/(text)/i.test(cs.backgroundClip || cs.webkitBackgroundClip || '') && cs.backgroundImage && cs.backgroundImage !== 'none') {
			const colors = String(cs.backgroundImage).match(/rgba?\([^)]+\)|#[0-9a-fA-F]{3,8}/g)
			if (colors && colors.length) return colorToHex(colors[0], colorToHex(cs.color, '000000'))
		}
		return colorToHex(cs.color, '000000')
	}

	function inlineRuns(el, base) {
		const runs = []
		const walk = (node, inherited) => {
			if (node.nodeType === Node.TEXT_NODE) {
				const owner = node.parentElement || el
				const cs = getComputedStyle(owner)
				const value = applyTextTransform(normalizeRunText(node.textContent, cs.whiteSpace), cs.textTransform)
				if (value.trim()) runs.push({ text: value, options: { ...inherited } })
				return
			}
			if (node.nodeType !== Node.ELEMENT_NODE) return
			if (node.tagName === 'BR') {
				if (runs.length) runs[runs.length - 1].options.breakLine = true
				return
			}
			const cs = getComputedStyle(node)
			const opts = {
				...inherited,
				fontFace: mapFont(cs.fontFamily, node.textContent),
				fontSize: Math.max(7, pt(cs.fontSize)),
				color: textPaint(cs, node),
				bold: inherited.bold || isBoldWeight(cs.fontWeight) || node.tagName === 'STRONG' || node.tagName === 'B',
				italic: inherited.italic || cs.fontStyle === 'italic' || node.tagName === 'EM' || node.tagName === 'I',
				underline: inherited.underline || String(cs.textDecoration || '').includes('underline') || node.tagName === 'U'
			}
			if (node.tagName === 'A' && node.href) opts.hyperlink = { url: node.href, tooltip: node.title || node.href }
			Array.from(node.childNodes).forEach((child) => walk(child, opts))
			if (node !== el && /^(block|flex|grid|list-item|table|flow-root)$/i.test(cs.display || '') && runs.length) {
				runs[runs.length - 1].options.breakLine = true
			}
		}
		Array.from(el.childNodes).forEach((child) => walk(child, { ...(base || {}) }))
		if (!runs.length) runs.push({ text: text(el), options: { ...(base || {}) } })
		if (base && base.breakLine && runs.length) runs[runs.length - 1].options.breakLine = true
		return runs.filter((run) => run.text)
	}

	function normalizeRunText(value, whiteSpace) {
		const s = String(value || '').replace(/\u00a0/g, ' ')
		if (/pre/.test(whiteSpace || '')) return s.replace(/[ \t]+\n/g, '\n')
		return s.replace(/\s+/g, ' ')
	}

	function mark(el, descendants) {
		processed.add(el)
		if (descendants !== false) {
			el.querySelectorAll('*').forEach((node) => {
				// Never swallow icon/graphic nodes: they have their own handlers
				// (addInlineSvg / addShadowSvg) and must still be emitted even when an
				// ancestor decorative container is marked. <iconify-icon> is the
				// non-rasterized icon element and was being dropped here.
				const tag = node.tagName && node.tagName.toLowerCase()
				if (tag === 'svg' || tag === 'iconify-icon') return
				processed.add(node)
			})
		}
		return true
	}

	extractBackground()
	collectNotes()
	const handlers = [addShadowSvg, addInlineSvg, addImage, addTable, addRule, addContainer, addList, addText]
	Array.from(document.querySelectorAll('*')).forEach((el) => {
		if (processed.has(el)) return
		if (['SCRIPT', 'STYLE', 'LINK', 'META', 'TITLE', 'SOURCE', 'TRACK', 'PARAM', 'IFRAME', 'VIDEO', 'AUDIO', 'CANVAS', 'OBJECT', 'EMBED'].includes(el.tagName)) {
			mark(el)
			return
		}
		for (const handler of handlers) {
			if (handler(el)) return
		}
	})
	await normalizeImageSources()
	return model
}
