import type { Theme } from 'vitepress'
import DefaultTheme from 'vitepress/theme'
import { defineAsyncComponent, h } from 'vue'
import './wide.css'

/**
 * The default theme, extended in one place.
 *
 * Everything the site shows about a deck is produced by running the library —
 * there are no slide images to theme around, on purpose (see docs/decisions).
 * So this file exists to register components that *run* the loop, and nothing
 * else; restyling belongs in CSS, not here.
 *
 * `Playground` is registered asynchronously, which is what keeps the library off
 * every other page: the chunk holding `pptx-html` and `@shbernal/ts-pptx` is
 * fetched when the component mounts, and the component only mounts inside
 * `<ClientOnly>` on the playground page. A static import here would put a
 * megabyte of deck-writing code in the bundle a reader of the architecture page
 * downloads — and would also break the build outright, since VitePress
 * pre-renders every page in Node and the loop wants a DOM.
 */
export default {
	extends: DefaultTheme,
	enhanceApp({ app }) {
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
	},
} satisfies Theme
