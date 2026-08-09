import { playwright } from '@vitest/browser-playwright'
import { defineConfig } from 'vitest/config'

// Three test layers, kept as separate projects so each runs in the right
// environment:
//   - `unit`    — node, no DOM, fast. The pure pieces of both lanes: the paint
//                 model and its editable surface, the island and reconcile logic,
//                 and hand-written `heuristic/*` fixtures emitted and read back
//                 with ts-pptx's `read` model.
//   - `oracle`  — node. The round-trip gate: generates the corpus, runs each deck
//                 through a lane, and diffs the result under the normalized read
//                 model. The whole loop runs here — `parseDeck` needs no DOM, so
//                 the browser project only has to prove the DOM reading itself.
//   - `browser` — real Chromium via Playwright. Two things that cannot be faked:
//                 reading the editable surface back out of a rendered document,
//                 and the heuristic lane's extract path (`convertSlide` /
//                 `convertDeck`) with an injected offline `resolveIcon`.
//                 jsdom/happy-dom are insufficient (no real
//                 getBoundingClientRect / getComputedStyle layout / canvas).
export default defineConfig({
	test: {
		projects: [
			{
				test: {
					name: 'unit',
					environment: 'node',
					include: ['test/unit/**/*.test.ts'],
				},
			},
			{
				test: {
					name: 'oracle',
					environment: 'node',
					include: ['test/oracle/**/*.test.ts'],
					// Generating a deck per corpus entry is write-bound, not CPU-bound.
					testTimeout: 30_000,
				},
			},
			{
				test: {
					name: 'browser',
					include: ['test/browser/**/*.test.ts'],
					browser: {
						enabled: true,
						provider: playwright(),
						headless: true,
						screenshotFailures: false,
						instances: [{ browser: 'chromium' }],
					},
				},
			},
		],
	},
})
