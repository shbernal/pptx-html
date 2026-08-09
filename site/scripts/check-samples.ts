/**
 * Fails the site build when the playground's curated deck names have drifted from
 * the corpus.
 *
 * `samples()` throws on a missing name, but it throws in a visitor's browser,
 * which is far too late — the whole reason the playground draws from
 * `test/corpus/decks.ts` is that the decks it shows are the decks the oracle
 * gates, and a rename would otherwise turn that into a runtime error nobody sees
 * until someone opens the page. Running the same check here makes it a build
 * failure instead.
 */

import { samples } from '../.vitepress/theme/playground/samples.ts'

const found = samples()
console.log(`check-samples: ${found.length} corpus decks resolved (${found.map((s) => s.name).join(', ')})`)
