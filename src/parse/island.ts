/**
 * Finding the island in a document, and deciding whether to believe it.
 *
 * This is the first half of the return path, and it is deliberately the half
 * with no interpretation in it at all. It locates three `<script>` blocks by id,
 * checks that the model in them is one this build understands and has not been
 * altered, and hands back a {@link RenderIr}. Nothing here looks at the visual
 * DOM; nothing here guesses.
 *
 * ## Three outcomes, and only one of them is a lane
 *
 * - **absent** — no island. The document was written by something else (a hand-
 *   authored page, an agent's HTML), so the deck can only be inferred. That is
 *   the heuristic lane, and it is a legitimate, warned outcome.
 * - **present and sound** — the exact/reconciled lanes, decided later by
 *   comparing surfaces.
 * - **present and unsound** — a version this build cannot read, a block that is
 *   not JSON, or a hash mismatch. All three **throw**.
 *
 * The last one is worth being blunt about: a `modelHash` mismatch is not a lane
 * and must never become one. Falling back to inference on a tampered model would
 * take the one document that is known to be lying and hand it to the code path
 * with the weakest guarantees, which is precisely backwards. Nor is a version
 * mismatch something to migrate through: this build cannot verify a model whose
 * shape it does not know, so reading it optimistically would produce a deck
 * nobody checked.
 */

import { IR_VERSION, type RenderIr } from '../ir/render'
import { project } from '../ir/surface'
import { ASSETS_ID, type Integrity, INTEGRITY_ID, islandTextOf, ISLAND_ID, sha256Hex } from '../render/island'

/** Why a document that *has* an island cannot be read from it. */
export type IslandFault = 'malformed' | 'version' | 'tampered'

/**
 * A document carrying an island this build refuses to trust.
 *
 * A named class rather than a bare `Error` because the caller's response differs
 * by fault: `version` is "re-render from the source", `tampered` is "the file has
 * been altered", `malformed` is "this is not one of our documents after all".
 */
export class IslandError extends Error {
	// Declared and assigned rather than written as a constructor parameter
	// property: that syntax is TypeScript-only, and `erasableSyntaxOnly` keeps this
	// package to what a type-stripping runtime can erase.
	readonly fault: IslandFault

	constructor(fault: IslandFault, message: string) {
		super(message)
		this.fault = fault
		this.name = 'IslandError'
	}
}

/** The three blocks a rendered document may carry, as raw text. */
export interface DocumentBlocks {
	island: string | null
	integrity: string | null
	/** Present only under `assets: 'inline'`. */
	assets: string | null
}

/**
 * The text of one `<script type="application/json" id="…">` block.
 *
 * A string scanner rather than a parser, and that is safe for exactly one
 * reason: **every `<` inside a block this package writes is escaped** (see
 * `render/island.ts`), so the first `</script` after the opening tag is
 * guaranteed to be the block's own end. Without that guarantee this would be an
 * HTML-tokenizer bug waiting to happen; with it, the scan is exact.
 *
 * The DOM path below is preferred wherever a real document exists — this one is
 * what makes the return path work in Node, where the oracle runs and where no
 * `DOMParser` is available.
 */
function scanBlock(html: string, id: string): string | null {
	const marker = `id="${id}"`
	for (let from = 0; ; ) {
		const at = html.indexOf(marker, from)
		if (at < 0) return null
		from = at + marker.length

		const open = html.lastIndexOf('<script', at)
		if (open < 0) continue
		const tagEnd = html.indexOf('>', open)
		// The marker has to sit inside *this* tag; if the tag closed before it, the
		// match was text that happens to look like an id attribute.
		if (tagEnd < 0 || tagEnd < at) continue

		const close = html.indexOf('</script', tagEnd + 1)
		if (close < 0) return null
		return html.slice(tagEnd + 1, close)
	}
}

function domBlock(root: ParentNode, id: string): string | null {
	// `getElementById` is a `Document` method and this accepts any subtree root,
	// so the selector form is the one that works for both.
	return root.querySelector(`script#${id}`)?.textContent ?? null
}

/** Pull the blocks out of a document, however it arrived. */
export function blocksOf(source: string | ParentNode): DocumentBlocks {
	const read = typeof source === 'string' ? (id: string) => scanBlock(source, id) : (id: string) => domBlock(source, id)
	return { island: read(ISLAND_ID), integrity: read(INTEGRITY_ID), assets: read(ASSETS_ID) }
}

function parseJson(text: string, what: string): unknown {
	try {
		return JSON.parse(text)
	} catch (cause) {
		throw new IslandError('malformed', `the ${what} block is not JSON: ${(cause as Error).message}`)
	}
}

/** The island, verified. `null` means there was none — the heuristic lane. */
export interface OpenedIsland {
	ir: RenderIr
	/** The hashes as the document stated them, after they were checked. */
	integrity: Integrity
	/** The island block's text, verbatim. Kept because it is what `modelHash` covers. */
	text: string
}

/**
 * Verify and deserialize.
 *
 * The order matters: version before hash, because a model from a future renderer
 * would fail the hash check for an uninteresting reason and report "tampered" for
 * what is really "too new". Hash before use, always.
 *
 * The `surfaceHash` recheck at the end is not redundant with the `modelHash` one.
 * Both are derived from the same bytes, so an island that passes the first can
 * only fail the second if the *integrity block itself* was edited — a document
 * whose two halves disagree about what it contains. That is not a lane either.
 */
export async function openIsland(blocks: DocumentBlocks): Promise<OpenedIsland | null> {
	if (blocks.island === null) return null
	if (blocks.integrity === null) {
		// An island with no integrity block was not written by this renderer. Reading
		// it anyway would mean trusting a model with nothing to check it against,
		// which is the whole thing the two hashes exist to prevent.
		throw new IslandError('malformed', `the document has an ${ISLAND_ID} block but no ${INTEGRITY_ID} block`)
	}

	const stated = parseJson(blocks.integrity, 'integrity') as Partial<Integrity>
	if (typeof stated.irVersion !== 'number' || typeof stated.modelHash !== 'string' || typeof stated.surfaceHash !== 'string') {
		throw new IslandError('malformed', 'the integrity block is missing irVersion, modelHash or surfaceHash')
	}
	if (stated.irVersion !== IR_VERSION) {
		throw new IslandError(
			'version',
			`the document carries IR version ${stated.irVersion} and this build reads ${IR_VERSION}; re-render it from the source package rather than migrating a model that cannot be verified`
		)
	}

	const modelHash = await sha256Hex(blocks.island)
	if (modelHash !== stated.modelHash) {
		throw new IslandError(
			'tampered',
			`the island does not match its stated modelHash (${stated.modelHash.slice(0, 12)}… vs ${modelHash.slice(0, 12)}…); the model has been altered and nothing derived from it can be trusted`
		)
	}

	const ir = parseJson(blocks.island, 'island') as RenderIr
	if (ir.irVersion !== stated.irVersion) {
		throw new IslandError('malformed', `the island states IR version ${ir.irVersion} and the integrity block states ${stated.irVersion}`)
	}

	const surfaceHash = await sha256Hex(JSON.stringify(project(ir)))
	if (surfaceHash !== stated.surfaceHash) {
		throw new IslandError(
			'tampered',
			'the island and the integrity block disagree about the editable surface; the integrity block has been altered'
		)
	}

	return { ir, integrity: { irVersion: stated.irVersion, modelHash, surfaceHash }, text: blocks.island }
}

/**
 * The surface hash of an arbitrary IR, in the one spelling both sides use.
 *
 * Here rather than inlined at each call site so the renderer's hash and the
 * parser's comparison can never be taken over differently-serialized forms of the
 * same projection — the failure that would report every untouched document as
 * edited.
 */
export async function surfaceHashOf(ir: RenderIr): Promise<string> {
	return sha256Hex(JSON.stringify(project(ir)))
}

/**
 * Re-derive `modelHash` for an IR, for a caller checking a package against a
 * document. Over {@link islandTextOf}, because that is the form the hash covers.
 */
export async function modelHashOf(ir: RenderIr): Promise<string> {
	return sha256Hex(islandTextOf(ir))
}
