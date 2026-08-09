/**
 * Single-file components are compiled by Vite, not by `tsc`, so TypeScript needs
 * telling that a `.vue` import resolves to a component. Deliberately loose: the
 * playground's logic lives in plain `.ts` beside these files precisely so the
 * part worth checking is checked, and a stricter shim here would only be
 * pretending to check the templates.
 */
declare module '*.vue' {
	import type { DefineComponent } from 'vue'

	const component: DefineComponent<Record<string, unknown>, Record<string, unknown>, unknown>
	export default component
}

/** A stylesheet imported for its side effect. Vite injects it; there is nothing to type. */
declare module '*.css' {}
