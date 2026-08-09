import type { Theme } from 'vitepress'
import DefaultTheme from 'vitepress/theme'
import { defineAsyncComponent, h } from 'vue'
import HomeHero from './home/HomeHero.vue'
import './home/home.css'
import './wide.css'

/**
 * The default theme, extended in one place.
 *
 * Everything the site shows about a deck is produced by running the library —
 * there are no slide images to theme around, on purpose (see docs/decisions).
 * So this file exists to register components that *run* the loop, plus the one
 * that does not; restyling belongs in CSS, not here.
 *
 * `Playground` and `SlideMarquee` are registered asynchronously, which is what
 * keeps the library off every other page: the chunk holding `pptx-html` and
 * `@shbernal/ts-pptx` is fetched when one of them mounts, and both only mount
 * inside `<ClientOnly>` — the playground on its page, the marquee on the home
 * page. A static import here would put a megabyte of deck-writing code in the
 * bundle a reader of the architecture page downloads, and would also break the
 * build outright, since VitePress pre-renders every page in Node and the loop
 * wants a DOM.
 *
 * `HomeHero` is the exception and is imported normally: it is static markup that
 * touches nothing, so it is pre-rendered into the home page's HTML instead of
 * appearing once Vue and a megabyte of writer have arrived.
 */
export default {
	extends: DefaultTheme,
	enhanceApp({ app }) {
		app.component('HomeHero', HomeHero)

		app.component(
			'Playground',
			defineAsyncComponent({
				loader: () => import('./playground/Playground.vue'),
				loadingComponent: {
					render: () => h('p', { class: 'pxh-booting' }, 'Loading pptx-html and the writer under it…'),
				},
				errorComponent: {
					render: () =>
						h(
							'p',
							{ class: 'pxh-booting' },
							'The playground failed to load. Nothing here falls back to a picture of a deck, so there is nothing to show in its place.'
						),
				},
			})
		)

		app.component(
			'SlideMarquee',
			defineAsyncComponent({
				loader: () => import('./home/SlideMarquee.vue'),
				loadingComponent: {
					render: () => h('p', { class: 'pxh-booting' }, 'Loading the writer and the renderer…'),
				},
				errorComponent: {
					render: () =>
						h(
							'p',
							{ class: 'pxh-booting' },
							'The showcase failed to load, so this space is empty. Nothing on this site stands in for a slide it did not draw.'
						),
				},
			})
		)
	},
} satisfies Theme
