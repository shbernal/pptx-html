/**
 * The heuristic lane's public surface, imported by **package name**.
 *
 * Not a test — nothing here runs. It exists to be typechecked, and the import
 * specifier is the whole point: `pptx-html` resolves through the `exports` map to
 * `dist/index.d.ts`, the way a consumer's editor resolves it, rather than through
 * a relative path into `src/`. A rename that breaks the entry point is invisible
 * to every other project in the repo, which is why `AGENTS.md` has to ask for
 * `dist/index.d.ts` to be *read* after touching `src/index.ts`.
 *
 * `examples/round-trip.mjs` covers the other half: it imports the same map at run
 * time, where a missing *value* export is a link error. The two are complements —
 * a runtime import cannot see a type-only export, and a type import cannot notice
 * that the value behind it stopped being emitted.
 *
 * The loop's own types are already exercised this way by `site/`, which imports
 * `pptx-html` for the playground. Nothing did the same for the lane below.
 *
 * This file needs `dist/`, so `pnpm run typecheck` needs a build first — which it
 * already did for `site/tsconfig.json`, for the same reason.
 */

import type {
	ConvertOptions,
	ConvertResult,
	heuristic,
	IconResolver,
	OutputMode,
	ProgressEvent,
	Warning,
} from 'pptx-html'
import { convertDeck, convertSlide } from 'pptx-html'

/** Every option the lane documents, spelled out rather than inferred. */
export const OPTIONS: ConvertOptions = {
	author: 'pptx-html',
	title: 'public surface',
	lang: 'en',
	fileName: 'deck.pptx',
	output: 'base64' satisfies OutputMode,
	vectorizeSvg: true,
	onProgress: (event: ProgressEvent) => void event.phase,
	shouldStop: () => false,
	resolveIcon: (async () => null) satisfies IconResolver,
}

/** The two entry points, at their declared signatures. */
export const DECK: (html: string) => Promise<ConvertResult> = (html) => convertDeck(html, OPTIONS)
export const SLIDE: (head: string, slide: string) => Promise<{ model: heuristic.SlideModel; warnings: Warning[] }> = (
	head,
	slide
) => convertSlide(head, slide, OPTIONS)
