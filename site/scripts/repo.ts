/**
 * Repository facts the site needs, read from the files that already own them.
 *
 * Nothing here may be hardcoded a second time in `config.ts` or in a markdown
 * page: the repository URL and the package description live in `package.json`,
 * and a copy of either would be right on the day it was written and wrong later.
 */

import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

/** Resolved from this file, not from `process.cwd()`, so the scripts run from anywhere. */
export const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..', '..')

interface PackageJson {
	name: string
	description: string
	repository: { url: string }
}

export const pkg: PackageJson = JSON.parse(readFileSync(resolve(repoRoot, 'package.json'), 'utf8'))

/** `git+https://github.com/owner/repo.git` → `https://github.com/owner/repo` */
export const repoUrl = pkg.repository.url.replace(/^git\+/, '').replace(/\.git$/, '')

/** A link to a file that lives in the repository and is deliberately not published on the site. */
export function blobUrl(path: string): string {
	return `${repoUrl}/blob/main/${path}`
}
