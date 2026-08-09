import { defineConfig } from 'vitepress'
import { LEDGER, readDocsSource } from '../scripts/docs-source.ts'
import { pkg, repoUrl } from '../scripts/repo.ts'

// Reading `docs/docs.json` here rather than restating the nav means the sidebar
// cannot drift from the design record: a page added to one and not the other
// fails the config, not a reader.
const docs = readDocsSource()

/** `index` is served at the directory root; every other page at its slug. */
function docLink(slug: string): string {
	return slug === 'index' ? '/docs/' : `/docs/${slug}`
}

export default defineConfig({
	// GitHub Pages serves the repo at https://shbernal.github.io/pptx-html/. Every
	// URL on the site goes through this, so nothing may hardcode a leading `/`.
	base: '/pptx-html/',
	title: 'pptx-html',
	description: pkg.description,
	lang: 'en-GB',

	// Left on deliberately: it is the acceptance test for the link rewriting in
	// site/scripts/docs-source.ts.
	ignoreDeadLinks: false,

	themeConfig: {
		nav: [
			{ text: 'Playground', link: '/playground' },
			{ text: 'Docs', link: '/docs/' },
			{ text: 'Changelog', link: `${repoUrl}/blob/main/CHANGELOG.md` },
		],

		sidebar: {
			'/docs/': [
				...docs.groups.map((group) => ({
					text: group.text,
					items: group.pages.map((page) => ({ text: page.title, link: docLink(page.slug) })),
				})),
				// Appended rather than added to `docs.json`: the ledger is generated from
				// the corpus, so listing it there would put a file that does not exist in
				// `docs/` into the record's own table of contents.
				{ text: 'Evidence', items: [{ text: LEDGER.title, link: docLink(LEDGER.slug) }] },
			],
		},

		search: { provider: 'local' },
		socialLinks: [{ icon: 'github', link: repoUrl }],

		// `:path` is relative to srcDir (`site/`), so a docs page resolves to
		// `docs/<slug>.md` — the tracked file, not the generated mirror it was
		// rendered from. Pages outside `docs/` set `editLink: false` in their
		// frontmatter, because there is no repository file behind them.
		editLink: {
			pattern: `${repoUrl}/edit/main/:path`,
			text: 'Edit this page on GitHub',
		},

		footer: {
			message: 'Released under the MIT License.',
			copyright: 'Copyright © shbernal',
		},
	},
})
