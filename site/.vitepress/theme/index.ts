import type { Theme } from 'vitepress'
import DefaultTheme from 'vitepress/theme'

/**
 * The default theme, extended in one place.
 *
 * Everything the site shows about a deck is produced by running the library —
 * there are no slide images to theme around, on purpose (see docs/decisions).
 * So this file exists to register components that *run* the loop, and nothing
 * else; restyling belongs in CSS, not here.
 */
export default {
	extends: DefaultTheme,
} satisfies Theme
