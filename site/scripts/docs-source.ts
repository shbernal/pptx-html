/**
 * Reads `docs/` — the tracked design record — as the site's input.
 *
 * `docs/` is written for the repository: its frontmatter carries keys the repo's
 * own docs tooling consumes, and its links point at files (`README.md`,
 * `CONTRIBUTING.md`) that the site does not publish. This module translates that
 * into what VitePress needs, and **fails rather than guesses** at every step. A
 * page the nav does not list, a frontmatter key nobody has decided about, a link
 * leaving `docs/` with no mapping — each one throws. The value of the mirror is
 * precisely that `docs/` cannot rot behind the site, and it can only deliver that
 * if drift is an error instead of a warning.
 */

import { readdirSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { blobUrl, repoRoot } from './repo.ts'

const DOCS_DIR = resolve(repoRoot, 'docs')

/**
 * The one page under `/docs/` that has no file in `docs/`.
 *
 * It is measured, not written: `site/scripts/build-ledger.ts` runs the oracle's
 * coverage reporter over the corpus and emits it into the mirror. Named here so
 * the generator and the sidebar agree, and kept out of `docs/docs.json` on
 * purpose — that file describes the tracked design record, and listing a page
 * that does not exist in `docs/` would break the cross-check that makes the
 * record and the site impossible to drift apart.
 */
export const LEDGER = {
	slug: 'fidelity',
	title: 'Fidelity ledger',
	description: 'What the loop models, carries and warns on, per corpus deck: measured, not asserted.',
}

/** Frontmatter the site renders: repo key → VitePress key. */
const CARRIED: Record<string, string> = { title: 'title', summary: 'description' }

/** Frontmatter that belongs to the repo's docs tooling and must not reach VitePress. */
const DROPPED = new Set(['doc-schema-version', 'read_when', 'doc_type'])

export interface DocPage {
	/** File stem, which is also the page's URL segment: `round-trip`. */
	slug: string
	title: string
	description: string
	/** Markdown after the frontmatter block, with links already rewritten. */
	body: string
}

export interface DocsSource {
	/** Nav groups in `docs.json` order. */
	groups: { text: string; pages: DocPage[] }[]
	pages: DocPage[]
}

interface DocsJson {
	navigation: { group: string; pages: string[] }[]
}

export function readDocsSource(): DocsSource {
	const nav: DocsJson = JSON.parse(readFileSync(resolve(DOCS_DIR, 'docs.json'), 'utf8'))
	const slugs = readdirSync(DOCS_DIR)
		.filter((name) => name.endsWith('.md'))
		.map((name) => name.slice(0, -3))
		.sort()
	const listed = nav.navigation.flatMap((group) => group.pages)

	const duplicated = listed.filter((slug, i) => listed.indexOf(slug) !== i)
	if (duplicated.length > 0) {
		throw new Error(`docs/docs.json lists the same page twice: ${duplicated.join(', ')}`)
	}
	const orphanNav = listed.filter((slug) => !slugs.includes(slug))
	if (orphanNav.length > 0) {
		throw new Error(`docs/docs.json navigates to pages with no file: ${orphanNav.join(', ')}`)
	}
	const orphanFiles = slugs.filter((slug) => !listed.includes(slug))
	if (orphanFiles.length > 0) {
		throw new Error(
			`docs/ holds pages the nav does not list: ${orphanFiles.join(', ')}. ` +
				'Add them to docs/docs.json, or the site publishes a page nothing links to.'
		)
	}

	const byslug = new Map(slugs.map((slug) => [slug, readPage(slug, slugs)]))
	return {
		groups: nav.navigation.map((group) => ({
			text: group.group,
			pages: group.pages.map((slug) => lookup(byslug, slug)),
		})),
		pages: slugs.map((slug) => lookup(byslug, slug)),
	}
}

function lookup(pages: Map<string, DocPage>, slug: string): DocPage {
	const page = pages.get(slug)
	// Reaches whoever broke the build, never a reader of the site.
	// charcheck-disable-next-line no-em-dash-in-site-build-strings
	if (!page) throw new Error(`docs/${slug}.md was validated and then not found — this is a bug in docs-source.ts`)
	return page
}

function readPage(slug: string, slugs: string[]): DocPage {
	const file = `docs/${slug}.md`
	const { data, body, bodyLine } = parseFrontmatter(readFileSync(resolve(DOCS_DIR, `${slug}.md`), 'utf8'), file)

	for (const key of data.keys()) {
		if (!(key in CARRIED) && !DROPPED.has(key)) {
			throw new Error(
				`${file}: unknown frontmatter key "${key}". Decide in site/scripts/docs-source.ts whether the ` +
					// Reaches whoever broke the build, never a reader of the site.
					// charcheck-disable-next-line no-em-dash-in-site-build-strings
					'site carries it or drops it — silently dropping it would publish a page missing something.'
			)
		}
	}

	return {
		slug,
		title: scalar(data, 'title', file),
		description: scalar(data, 'summary', file),
		body: rewriteLinks(body, slugs, file, bodyLine),
	}
}

// ---------------------------------------------------------------------------
// Frontmatter
//
// A deliberately narrow YAML reader: the shapes `docs/` actually uses, and an
// error for everything else. Accepting more would mean guessing what a
// half-understood construct meant, and the guess would land on the public site.
// ---------------------------------------------------------------------------

type Frontmatter = Map<string, string | string[]>

/**
 * A capture group's text. No pattern in this file makes its groups optional, so
 * an absent one means the pattern and the code reading it have drifted apart.
 * That is worth failing on rather than papering over with an empty string.
 */
function captured(match: RegExpExecArray, index: number): string {
	const value = match[index]
	if (value === undefined) throw new Error(`pattern matched ${JSON.stringify(match[0])} but has no capture ${index}`)
	return value
}

const FRONTMATTER = /^---\r?\n([\s\S]*?)\r?\n---\r?\n/

function parseFrontmatter(source: string, file: string): { data: Frontmatter; body: string; bodyLine: number } {
	const match = FRONTMATTER.exec(source)
	if (!match) throw new Error(`${file}: no frontmatter block; every page in docs/ carries one.`)

	const data: Frontmatter = new Map()
	const lines = captured(match, 1).split(/\r?\n/)
	let key: string | undefined

	for (const [i, line] of lines.entries()) {
		// +2: the opening `---` is line 1, so frontmatter line `i` is file line `i + 2`.
		const where = `${file}:${i + 2}`
		if (line.trim() === '') continue

		const item = /^\s+-\s+(.*)$/.exec(line)
		if (item) {
			const list = key === undefined ? undefined : data.get(key)
			if (!Array.isArray(list)) throw new Error(`${where}: list item does not follow a key with an empty value.`)
			list.push(unquote(captured(item, 1)))
			continue
		}

		const pair = /^([A-Za-z][\w-]*):\s*(.*)$/.exec(line)
		if (!pair) throw new Error(`${where}: neither a "key: value" line nor a "  - item" line: ${JSON.stringify(line)}`)
		key = captured(pair, 1)
		if (data.has(key)) throw new Error(`${where}: duplicate frontmatter key "${key}".`)
		const value = captured(pair, 2)
		data.set(key, value === '' ? [] : unquote(value))
	}

	return { data, body: source.slice(match[0].length), bodyLine: lines.length + 3 }
}

function unquote(value: string): string {
	if (value.length >= 2 && value.startsWith("'") && value.endsWith("'")) {
		return value.slice(1, -1).replaceAll("''", "'")
	}
	if (value.length >= 2 && value.startsWith('"') && value.endsWith('"')) return value.slice(1, -1)
	return value
}

function scalar(data: Frontmatter, key: string, file: string): string {
	const value = data.get(key)
	if (typeof value !== 'string' || value === '') {
		throw new Error(`${file}: frontmatter "${key}" must be a non-empty string; the site renders it.`)
	}
	return value
}

// ---------------------------------------------------------------------------
// Links
// ---------------------------------------------------------------------------

const LINK = /\]\(([^()\s]+)((?:\s+"[^"]*")?)\)/g
const ABSOLUTE = /^(?:[a-z][a-z0-9+.-]*:|\/\/)/i
const FENCE = /^\s*(?:```|~~~)/
const SIBLING = /^(?:\.\/)?([\w-]+)\.md$/

/**
 * Rewrites the links `docs/` writes for a reader of the repository into links a
 * reader of the site can follow. Code fences are left alone — a link inside one
 * is sample text, not navigation.
 */
function rewriteLinks(body: string, slugs: string[], file: string, bodyLine: number): string {
	let fenced = false
	return body
		.split('\n')
		.map((line, i) => {
			if (FENCE.test(line)) {
				fenced = !fenced
				return line
			}
			if (fenced) return line
			return line.replace(LINK, (_whole, target: string, title: string) => {
				return `](${rewriteTarget(target, slugs, `${file}:${bodyLine + i}`)}${title})`
			})
		})
		.join('\n')
}

function rewriteTarget(target: string, slugs: string[], where: string): string {
	if (target.startsWith('#') || ABSOLUTE.test(target)) return target

	const hashAt = target.indexOf('#')
	const path = hashAt === -1 ? target : target.slice(0, hashAt)
	const hash = hashAt === -1 ? '' : target.slice(hashAt)

	const sibling = SIBLING.exec(path)
	const siblingSlug = sibling === null ? null : captured(sibling, 1)
	if (siblingSlug !== null && slugs.includes(siblingSlug)) return `./${siblingSlug}${hash}`

	// The landing page stands in for the README. A leading `/` is what VitePress
	// prefixes with `base`, so this stays correct under `/pptx-html/`.
	if (path === '../README.md') return `/${hash}`

	// CONTRIBUTING is for people working in a checkout; the site never publishes
	// it, so the honest destination is the file itself on GitHub.
	if (path === '../CONTRIBUTING.md') return `${blobUrl('CONTRIBUTING.md')}${hash}`

	throw new Error(
		// Reaches whoever broke the build, never a reader of the site.
		// charcheck-disable-next-line no-em-dash-in-site-build-strings
		`${where}: link "${target}" leaves docs/ and has no mapping. Add one in site/scripts/docs-source.ts — ` +
			'a guess here becomes a broken or misleading link on the public site.'
	)
}
