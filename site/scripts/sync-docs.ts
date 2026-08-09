/**
 * Writes `site/docs/` — the generated mirror of `docs/`.
 *
 * Runs before `vitepress dev` and before `vitepress build`. It rewrites the whole
 * directory every time rather than diffing: the mirror is cheap, and a stale file
 * in it is invisible until someone reads a page that quietly disagrees with the
 * repository.
 */

import { mkdirSync, rmSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { type DocPage, readDocsSource } from './docs-source.ts'
import { repoRoot } from './repo.ts'

const OUT_DIR = resolve(repoRoot, 'site', 'docs')

const { pages } = readDocsSource()

rmSync(OUT_DIR, { recursive: true, force: true })
mkdirSync(OUT_DIR, { recursive: true })
for (const page of pages) {
	writeFileSync(resolve(OUT_DIR, `${page.slug}.md`), render(page), 'utf8')
}

console.log(`sync-docs: ${pages.length} pages from docs/ → site/docs/`)

function render(page: DocPage): string {
	return [
		'---',
		`title: ${yaml(page.title)}`,
		`description: ${yaml(page.description)}`,
		'---',
		'',
		`<!-- Generated from docs/${page.slug}.md by site/scripts/sync-docs.ts. Edit the source, not this file. -->`,
		'',
		page.body.replace(/^\n+/, ''),
	].join('\n')
}

function yaml(value: string): string {
	return `'${value.replaceAll("'", "''")}'`
}
