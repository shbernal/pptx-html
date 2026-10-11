/**
 * A package reduced to a set of XML facts, and the diff that finds an option's
 * footprint in it.
 *
 * A probe writes the same deck twice, with and without one option, and the facts
 * that differ are what that option put in the package: its **signature**. Then it
 * asks whether the loop's output still holds that signature. Nobody writes down
 * what `underline: { style: 'dbl' }` should look like in OOXML; the writer says,
 * by writing it.
 *
 * This is deliberately *not* the oracle's equality. `diffDeckIr` compares two
 * `DeckIr`s, and a value the IR does not carry is missing from both sides alike
 * and compares clean (docs/round-trip.md, "What a green oracle does not prove").
 * The XML is the one view of the output that does not depend on what the read
 * model chose to model. Reading it here is test code looking at what the writer
 * produced; nothing in `src/` ever touches XML.
 *
 * A fact is `part-kind ⟩ element path @ attribute = value`, or `# = text` for
 * character data. Positions are dropped (a run's facts do not say *which* run),
 * which costs precision and buys robustness: the loop may reorder parts, renumber
 * them, and renumber ids, and none of that is a loss.
 *
 * Repeats are counted, though: the second identical fact in a part kind is the
 * fact `… ×2`, and so on. Without that, an option whose whole effect is *more of
 * the same* (`breakLine` splitting one paragraph into two alike) writes nothing a
 * set can see, and a loop that merged them back would lose nothing a set can see.
 *
 * A relationship is counted within its `Type`, since its `Id` and `Target` are
 * identities. Counted together, one relationship the loop drops (the empty notes
 * slide the writer gives every slide) would shift the count of every other, and a
 * probe that adds a link or a picture would lose `Relationship ×3` while its own
 * relationship survived.
 */

import { createHash } from 'node:crypto'
import { readZip } from 'pptx-ts/zip'

const decoder = new TextDecoder()

/**
 * Attributes whose values are identities rather than content. The loop is free to
 * renumber all of them, so they are reduced to `*`.
 */
const IDENTITY_ATTRIBUTES = new Set([
	'r:id',
	'r:embed',
	'r:link',
	'r:pict',
	'r:dm',
	'r:lo',
	'r:qs',
	'r:cs',
	'Id',
	'Target',
	'spid',
	'grpId',
	'id',
])

/** Parts that hold no slide content, or hold the clock. */
const IGNORED_PARTS = [
	/^docProps\//,
	/^\[Content_Types\]\.xml$/,
	/^ppt\/presProps\.xml$/,
	/^ppt\/viewProps\.xml$/,
	/^ppt\/tableStyles\.xml$/,
]

/**
 * `ppt/slides/slide3.xml` → `ppt/slides/slide#.xml`: parts are compared by kind, not by number.
 * A media part's name says nothing at all (the slide reaches it through a relationship
 * whose `Target` is already an identity), so it keeps only its extension:
 * `ppt/media/image-1-1.png` and `ppt/media/image1.png` are both `ppt/media/*.png`.
 */
function partKind(name: string): string {
	if (name.startsWith('ppt/media/')) return `ppt/media/*${/\.[^./]*$/.exec(name)?.[0] ?? ''}`
	return name.replace(/\d+/g, '#')
}

/** A binary part's content, short enough to read: two different images of one size are two facts. */
function binaryFact(data: Uint8Array): string {
	return `<binary ${data.byteLength} ${createHash('sha256').update(data).digest('hex').slice(0, 12)}>`
}

/** Adds the next occurrence of `fact`: itself the first time, `fact ×n` the n-th. */
function addCounted(fact: string, into: Map<string, number>): void {
	into.set(fact, (into.get(fact) ?? 0) + 1)
}

/** A relationship's `Type`, last segment only: `…/relationships/hyperlink` → `hyperlink`. */
function relationshipType(attributes: string): string {
	const match = /\bType\s*=\s*(?:"([^"]*)"|'([^']*)')/.exec(attributes)
	return (match?.[1] ?? match?.[2] ?? '').split('/').pop() ?? ''
}

/** Every fact in one XML document. */
function xmlFacts(kind: string, xml: string, into: Map<string, number>): void {
	const stack: string[] = []
	const tag =
		/<(\/?)([A-Za-z_][\w.:-]*)((?:\s+[^\s=>/]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>|<\?[\s\S]*?\?>|<!--[\s\S]*?-->|<!\[CDATA\[([\s\S]*?)\]\]>/g
	let last = 0
	for (const match of xml.matchAll(tag)) {
		const between = xml.slice(last, match.index).trim()
		last = match.index + match[0].length
		if (between && stack.length > 0) addCounted(`${kind}⟩${stack.join('/')}#=${between}`, into)
		const [, close, name, attributes, selfClose, cdata] = match
		if (cdata !== undefined && stack.length > 0) addCounted(`${kind}⟩${stack.join('/')}#=${cdata}`, into)
		if (!name) continue
		if (close) {
			stack.pop()
			continue
		}
		stack.push(name === 'Relationship' ? `Relationship[${relationshipType(attributes ?? '')}]` : name)
		const path = `${kind}⟩${stack.join('/')}`
		addCounted(path, into)
		for (const attribute of (attributes ?? '').matchAll(/([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) {
			const key = attribute[1] as string
			const value = IDENTITY_ATTRIBUTES.has(key) ? '*' : (attribute[2] ?? attribute[3] ?? '')
			if (key.startsWith('xmlns')) continue
			addCounted(`${path}@${key}=${value}`, into)
		}
		if (selfClose) stack.pop()
	}
}

/** Every fact in a package's XML parts. Binary parts contribute their presence, size and content hash. */
export async function packageFacts(bytes: Uint8Array): Promise<Set<string>> {
	const counts = new Map<string, number>()
	for (const [name, data] of await readZip(bytes)) {
		if (IGNORED_PARTS.some((pattern) => pattern.test(name))) continue
		const kind = partKind(name)
		if (/\.(xml|rels)$/i.test(name)) xmlFacts(kind, decoder.decode(data), counts)
		else addCounted(`${kind}⟩${binaryFact(data)}`, counts)
	}
	const facts = new Set<string>()
	for (const [fact, count] of counts) for (let n = 1; n <= count; n++) facts.add(n === 1 ? fact : `${fact} ×${n}`)
	return facts
}

/** What an option put into the package, and what it took out. */
export interface Signature {
	added: string[]
	removed: string[]
}

export function signature(baseline: Set<string>, probe: Set<string>): Signature {
	return {
		added: [...probe].filter((fact) => !baseline.has(fact)).toSorted(),
		removed: [...baseline].filter((fact) => !probe.has(fact)).toSorted(),
	}
}

/** The part of a signature an output failed to keep: added facts it lacks, removed facts it has again. */
export function lost(sig: Signature, output: Set<string>): string[] {
	return [
		...sig.added.filter((fact) => !output.has(fact)),
		...sig.removed.filter((fact) => output.has(fact)).map((fact) => `(reappeared) ${fact}`),
	]
}
