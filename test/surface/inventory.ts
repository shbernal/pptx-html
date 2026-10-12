/**
 * The writer's option surface, enumerated from its own type declarations.
 *
 * This is the denominator of the surface score: every option path the installed
 * `pptx-ts` lets an author set, e.g. `addText.underline.style` or
 * `addTable.cell.border.color`. It is read from the shipped `.d.ts` files rather
 * than written down here, so a release that adds an option grows the frontier on
 * its own and nobody has to notice.
 *
 * TypeScript 7 ships no compiler API, so this is a small structural reader over
 * declaration files the bundler emitted: top-level `interface` blocks and `type`
 * aliases, their members, `extends` (with `Omit<…>` and `Pick<…>` honoured), inline object types
 * and alias expansion. It does not need to understand TypeScript in general, only
 * the shape tsdown writes, and `surface.test.ts` pins that it still finds the
 * options the corpus is known to use.
 *
 * Self-contained on purpose (node built-ins only, explicit extensions): the score
 * script imports it under plain `node`, outside vitest.
 */

import { readdirSync, readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { dirname, join } from 'node:path'

export interface OptionPath {
	/** `addText.underline.style`: the authoring call, then the option keys. */
	path: string
	/** The authoring call the path starts from; the unit a coverage run works through. */
	family: string
	/** The declared type text, for the agent reading the frontier. */
	type: string
}

/** Where each family starts: the call an author makes and the type its options take. */
const ROOTS: Record<string, string> = {
	addText: 'TextPropsOptions',
	addShape: 'ShapeProps',
	addImage: 'ImageProps',
	addTable: 'TableProps',
	'addTable.cell': 'TableCellProps',
	addConnector: 'ConnectorProps',
	addGroup: 'GroupProps',
	addChart: 'ChartOpts',
	addNotes: 'NotesProps',
	addMedia: 'MediaProps',
	addComment: 'CommentProps',
	addAnimation: 'AnimationProps',
	slide: 'SlideSurface',
	defineSlideMaster: 'SlideMasterProps',
	theme: 'ThemeProps',
}

const ROOT_TYPES = new Set(Object.values(ROOTS))

/** `Slide`'s own settable properties, which no interface names on their own. */
const SYNTHETIC = `interface SlideSurface {
  transition?: TransitionProps;
  background?: BackgroundOption;
  color?: HexColor;
  hidden?: boolean;
  slideNumber?: SlideNumberProps;
}`

/** How far below the call option paths are followed; deeper ones are the call's own business. */
const MAX_DEPTH = 4

interface Member {
	name: string
	type: string
}

interface Declarations {
	interfaces: Map<string, { extendsText: string; members: Member[] }>
	aliases: Map<string, string>
}

function stripComments(source: string): string {
	return source.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
}

/** Index of the bracket closing the one at `open`, counting `{}`, `()`, `[]` and `<>`. */
function closing(text: string, open: number): number {
	let depth = 0
	for (let index = open; index < text.length; index++) {
		const char = text[index]
		if (char === '{' || char === '(' || char === '[' || char === '<') depth++
		else if (char === '}' || char === ')' || char === ']' || (char === '>' && text[index - 1] !== '=')) depth--
		if (depth === 0) return index
	}
	return text.length
}

/** Split a type body's members at top-level `;` and newlines. */
function membersOf(body: string): Member[] {
	const members: Member[] = []
	let depth = 0
	let start = 0
	const flush = (end: number) => {
		const chunk = body.slice(start, end).trim()
		start = end + 1
		const match = /^(?:readonly\s+)?['"]?([A-Za-z_$][\w$]*)['"]?\??\s*:\s*([\s\S]+)$/.exec(chunk)
		if (match?.[1] && match[2]) members.push({ name: match[1], type: match[2].trim().replace(/[;,]$/, '') })
	}
	for (let index = 0; index < body.length; index++) {
		const char = body[index]
		if (char === '{' || char === '(' || char === '[' || char === '<') depth++
		else if (char === '}' || char === ')' || char === ']' || (char === '>' && body[index - 1] !== '=')) depth--
		else if (depth === 0 && (char === ';' || char === '\n')) {
			// A newline only ends a member when the next line starts a new one.
			if (char === '\n' && !/^\s*(?:readonly\s+)?['"]?[A-Za-z_$][\w$]*['"]?\??\s*:/.test(body.slice(index + 1)))
				continue
			flush(index)
		}
	}
	flush(body.length)
	return members
}

function readDeclarations(sources: string[]): Declarations {
	const interfaces: Declarations['interfaces'] = new Map()
	const aliases: Declarations['aliases'] = new Map()
	for (const raw of sources) {
		const source = stripComments(raw)
		for (const match of source.matchAll(/^(?:export\s+)?(?:declare\s+)?interface\s+(\w+)(?:<[^{]*?>)?\s*([^{]*)\{/gm)) {
			const name = match[1] as string
			const open = match.index + match[0].length - 1
			const body = source.slice(open + 1, closing(source, open))
			if (!interfaces.has(name))
				interfaces.set(name, {
					extendsText: match[2]?.replace(/^extends\s+/, '').trim() ?? '',
					members: membersOf(body),
				})
		}
		for (const match of source.matchAll(/^(?:export\s+)?(?:declare\s+)?type\s+(\w+)(?:<[^=]*?>)?\s*=\s*/gm)) {
			const name = match[1] as string
			let end = match.index + match[0].length
			let depth = 0
			for (; end < source.length; end++) {
				const char = source[end]
				if (char === '{' || char === '(' || char === '[' || char === '<') depth++
				else if (char === '}' || char === ')' || char === ']' || (char === '>' && source[end - 1] !== '=')) depth--
				else if (depth === 0 && char === ';') break
			}
			if (!aliases.has(name)) aliases.set(name, source.slice(match.index + match[0].length, end).trim())
		}
	}
	return { interfaces, aliases }
}

/** `X['key']` replaced by the declared type of `key` on `X`, so the reader never walks all of `X`. */
function resolveIndexed(decls: Declarations, type: string): string {
	return type.replace(/\b([A-Z]\w*)\['(\w+)'\]/g, (whole, name: string, key: string) => {
		const member = objectMembers(decls, name, new Set()).find((candidate) => candidate.name === key)
		return member ? `(${member.type})` : whole
	})
}

/** The object members a type text resolves to, through aliases, unions, intersections and `extends`. */
function objectMembers(decls: Declarations, raw: string, seen: Set<string>): Member[] {
	const type = resolveIndexed(decls, raw)
	const members = new Map<string, Member>()
	const add = (list: Member[]) => {
		for (const member of list) if (!members.has(member.name)) members.set(member.name, member)
	}

	// Inline object literals, at any position in the text.
	for (let index = type.indexOf('{'); index !== -1;) {
		const end = closing(type, index)
		add(membersOf(type.slice(index + 1, end)))
		index = type.indexOf('{', end)
	}

	// `Omit<X, 'a' | 'b'>`: X's members minus the omitted keys.
	const omitted = new Set<string>()
	for (const match of type.matchAll(/Omit<\s*(\w+)\s*,\s*([^>]+)>/g)) {
		for (const key of match[2]?.matchAll(/'(\w+)'/g) ?? []) omitted.add(key[1] as string)
	}

	// `Pick<X, 'a' | 'b'>`: only the picked keys of X, and X is not walked whole below.
	const PICK = /Pick<\s*(\w+)\s*,\s*([^>]+)>/g
	for (const match of type.matchAll(PICK)) {
		const name = match[1] as string
		if (seen.has(name)) continue
		const picked = new Set([...(match[2]?.matchAll(/'(\w+)'/g) ?? [])].map((key) => key[1]))
		add(objectMembers(decls, name, seen).filter((member) => picked.has(member.name)))
	}

	const stripped = type.replace(/\{[\s\S]*\}/g, ' ').replace(PICK, ' ')
	for (const match of stripped.matchAll(/\b([A-Z]\w*)\b/g)) {
		const name = match[1] as string
		if (seen.has(name) || name === 'Omit' || name === 'Partial' || name === 'Array' || name === 'Record') continue
		const next = new Set(seen).add(name)
		const iface = decls.interfaces.get(name)
		if (iface) {
			add(iface.members)
			if (iface.extendsText) add(objectMembers(decls, iface.extendsText, next))
			continue
		}
		const alias = decls.aliases.get(name)
		if (alias) add(objectMembers(decls, alias, next))
	}
	for (const key of omitted) members.delete(key)
	return [...members.values()]
}

/** Whether a type admits a non-object value (`string`, a literal, a number), through aliases. */
function hasScalar(decls: Declarations, raw: string, seen: Set<string>): boolean {
	const stripped = resolveIndexed(decls, raw).replace(/\{[\s\S]*\}/g, ' ')
	if (/\b(?:string|number|boolean|true|false)\b|'[^']*'|"[^"]*"|\d/.test(stripped.replace(/<[^>]*>/g, ' '))) return true
	for (const match of stripped.matchAll(/\b([A-Z]\w*)\b/g)) {
		const name = match[1] as string
		if (seen.has(name)) continue
		const alias = decls.aliases.get(name)
		if (alias && hasScalar(decls, alias, new Set(seen).add(name))) return true
	}
	return false
}

/** Type names a type text mentions; the ancestors a nested walk must not re-enter. */
function namesIn(type: string): string[] {
	return [...type.matchAll(/\b([A-Z]\w*)\b/g)].map((match) => match[1] as string)
}

function walk(
	decls: Declarations,
	family: string,
	prefix: string,
	type: string,
	ancestors: Set<string>,
	out: OptionPath[]
): void {
	const inner = new Set([...ancestors, ...namesIn(type)])
	const depth = prefix.split('.').length - family.split('.').length
	for (const member of objectMembers(decls, type, ancestors)) {
		const path = `${prefix}.${member.name}`
		// An option whose type is another family's whole option set (a master's text
		// objects, a table's per-cell options) is that family's surface, counted once
		// there. It stays one path here, so placing the object is still covered.
		const borrowed = namesIn(member.type).find((name) => ROOT_TYPES.has(name) && name !== ROOTS[family])
		if (borrowed) {
			out.push({ path, family, type: member.type.replace(/\s+/g, ' ') })
			continue
		}
		const children = depth < MAX_DEPTH ? objectMembers(decls, member.type, inner) : []
		// A path is a unit of coverage when it can be set to a value on its own. A
		// pure object option (`line: ShapeLineProps`) is covered by its children.
		if (children.length === 0 || hasScalar(decls, member.type, new Set()))
			out.push({ path, family, type: member.type.replace(/\s+/g, ' ') })
		if (children.length > 0) walk(decls, family, path, member.type, inner, out)
	}
}

/** The installed writer's `dist/` directory. */
function writerDist(): string {
	const require = createRequire(import.meta.url)
	return join(dirname(require.resolve('pptx-ts/package.json')), 'dist')
}

/** Every option path the installed writer accepts, sorted, families in {@link ROOTS} order. */
export function inventory(dist = writerDist()): OptionPath[] {
	const sources = readdirSync(dist)
		.filter((file) => file.endsWith('.d.ts'))
		.toSorted()
		.map((file) => readFileSync(join(dist, file), 'utf8'))
	const decls = readDeclarations([SYNTHETIC, ...sources])
	const out: OptionPath[] = []
	for (const [family, root] of Object.entries(ROOTS)) {
		const found: OptionPath[] = []
		walk(decls, family, family, root, new Set(), found)
		out.push(...found.toSorted((a, b) => a.path.localeCompare(b.path)))
	}
	return out
}

export const FAMILIES = Object.keys(ROOTS)
