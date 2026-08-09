/**
 * Deck/slide HTML parsing: split a full deck document into head + per-slide
 * sections, and compose a single-slide document for the render iframe.
 */

import { escapeAttr } from './frame'

/** A deck document split into the parts each render frame needs. */
export interface ParsedDeckHtml {
	headHTML: string
	slides: string[]
	lang: string
}

export function parseDeckHtml(fullHtml: unknown): ParsedDeckHtml {
	const doc = new DOMParser().parseFromString(String(fullHtml || ''), 'text/html')
	const rawHead = doc.head ? doc.head.innerHTML : ''
	const headHTML = rawHead
		.replace(/<script[\s\S]*?<\/script>/gi, (script) => (/iconify|tailwind|material/i.test(script) ? script : ''))
		.replace(/<link[^>]+fonts\.googleapis\.com[^>]*>/gi, '')
	let slides = Array.from(doc.querySelectorAll('.slide'))
		.filter((el) => !el.parentElement?.closest('.slide'))
		.map((el) => el.outerHTML)
	if (!slides.length) {
		slides = Array.from(doc.body.children)
			.filter((el) => /^(section|div|article)$/i.test(el.tagName))
			.map((el) => {
				el.classList.add('slide')
				return el.outerHTML
			})
	}
	const lang = doc.documentElement?.getAttribute('lang') || 'en'
	return { headHTML, slides, lang }
}

export function composeSlideDocument(headHTML: string, slideHTML: string, lang?: string): string {
	const exportCss =
		'<style>body.magic-exporting .slide,body.magic-exporting .frag,body.exporting .slide,body.exporting .frag{transition:none!important;animation:none!important;opacity:1!important;visibility:visible!important;transform:none!important}</style>'
	return (
		'<!doctype html><html lang="' +
		escapeAttr(lang || 'en') +
		'"><head><meta charset="utf-8">' +
		headHTML +
		exportCss +
		'</head><body class="exporting magic-exporting">' +
		slideHTML +
		'</body></html>'
	)
}
