import { playwright } from '@vitest/browser-playwright'
import { defineConfig } from 'vitest/config'

// Two test layers, kept as separate projects so each runs in the right
// environment:
//   - `unit`    — node, no DOM, fast. Feeds hand-written SlideModel fixtures into
//                 `emit/*` and parses the result back with ts-pptx's `read` model.
//   - `browser` — real Chromium via Playwright. Loads HTML slide fixtures through
//                 the full extract path (`convertSlide` / `convertDeck`) with an
//                 injected offline `resolveIcon`. jsdom/happy-dom are insufficient
//                 (no real getBoundingClientRect / getComputedStyle layout / canvas).
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
