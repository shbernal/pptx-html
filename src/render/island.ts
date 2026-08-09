/**
 * The IR island — the channel the return path trusts.
 *
 * The rendered document has two channels (see `document.ts`): a visual one made
 * of SVG and HTML, and this one, a verbatim copy of the {@link RenderIr} in a
 * `<script type="application/json">` block. Part 06 parses *this* and never
 * re-derives the model from the DOM, because `getComputedStyle` has no
 * representation for placeholder inheritance, colour transforms, autofit mode or
 * geometry adjust values — it would answer every question with a plausible
 * number and no way to tell which ones were invented.
 *
 * ## Why one blob and not `data-` attributes
 *
 * The model has to survive a DOM the browser has already normalized, and a
 * sanitizer or an editor that rewrites attributes. One opaque blob either
 * survives intact or is visibly gone; a model scattered across a hundred
 * `data-ir-*` attributes degrades silently, one attribute at a time. Per-element
 * `data-d2p-node` attributes still exist on the visual nodes, but only to *link*
 * a painted element back to its island entry — never to carry state.
 *
 * ## Two hashes, not one
 *
 * This is the part that is easy to get wrong by economising:
 *
 * - {@link Integrity.modelHash} covers the island's own bytes. It answers "has
 *   the model been tampered with", and any mismatch is a hard error.
 * - {@link Integrity.surfaceHash} covers `project(ir)` — the sanctioned editable
 *   projection and nothing else. It answers "did a human edit the text", and a
 *   mismatch is the normal signal that routes a slide to the reconciled lane.
 *
 * With a single hash those two questions collapse into one answer, and part 06
 * cannot tell a legitimate edit from corruption: every edit would look like
 * tampering, or nothing would. They are complementary by construction because
 * both derive from `src/ir/surface.ts` rather than from two separate ideas of
 * what is editable.
 */

import type { RenderIr } from '../ir/render'
import { project } from '../ir/surface'

/** The island block's `id`. Part 06 looks the model up by exactly this. */
export const ISLAND_ID = 'dom2pptx-ir'

/** The integrity block's `id` — the two hashes and the IR version. */
export const INTEGRITY_ID = 'dom2pptx-integrity'

/** The asset block's `id`, present only under `assets: 'inline'`. */
export const ASSETS_ID = 'dom2pptx-assets'

/**
 * What the integrity block holds.
 *
 * It is a *separate* block from the island rather than a field inside it, and
 * that is forced rather than stylistic: `modelHash` is taken over the island's
 * serialized bytes, so a copy of it living inside those bytes could never be
 * computed — the hash would have to contain itself.
 */
export interface Integrity {
	/**
	 * Repeated from the island so a reader can reject a document from an older
	 * renderer *before* parsing a model whose shape it does not know.
	 */
	irVersion: number
	/** Lowercase hex SHA-256 of the island's serialized JSON, as embedded. */
	modelHash: string
	/** Lowercase hex SHA-256 of the serialized {@link project} of the same IR. */
	surfaceHash: string
}

/**
 * Serialize the IR for embedding.
 *
 * No pretty-printing and no key reordering. The bytes written here are the exact
 * bytes {@link modelHashOf} covers and the exact bytes part 06 re-hashes, so any
 * transformation applied on one side and not the other turns every document into
 * a tampering report. `RenderIr` is JSON-safe by contract (no `undefined`, no
 * `Map`, no `Uint8Array` — see `ir/render`), so this needs no replacer.
 */
export function serializeIsland(ir: RenderIr): string {
	return JSON.stringify(ir)
}

/**
 * Make a JSON string safe to sit inside a `<script>` element.
 *
 * The HTML tokenizer ends a script block at the first `</script` it sees, even
 * inside what JSON considers a string — so a deck containing the literal text
 * `</script>` (an alt text, a run of text, a residual slide's XML) would
 * otherwise truncate the island and take the rest of the document with it.
 * Escaping `<` closes that off, and `<` is a JSON escape for the same
 * character, so the parsed value is unchanged. `<` cannot occur outside a string
 * in JSON, which is what makes the blanket replacement safe.
 *
 * This also disarms `<!--`, the other sequence that shifts the tokenizer's state
 * inside a script element.
 */
export function escapeForScript(json: string): string {
	return json.replaceAll('<', '\\u003c')
}

/** Lowercase hex SHA-256 of a string's UTF-8 bytes. */
export async function sha256Hex(text: string): Promise<string> {
	const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text))
	return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('')
}

/**
 * Both hashes for one IR, plus the version.
 *
 * `island` is passed in rather than re-serialized here so that the value hashed
 * is provably the value embedded — recomputing it would leave room for the two
 * to drift apart under a future change to {@link serializeIsland}.
 *
 * Neither hash covers asset *bytes*, only the manifest inside the island. That
 * is what makes `modelHash` identical under `assets: 'inline'` and
 * `assets: 'ref'`: the same model, delivered two ways, is one model. The bytes
 * are checked separately, against each manifest entry's own `sha256`, when they
 * are resolved.
 */
export async function integrityOf(ir: RenderIr, island: string): Promise<Integrity> {
	return {
		irVersion: ir.irVersion,
		modelHash: await sha256Hex(island),
		surfaceHash: await sha256Hex(JSON.stringify(project(ir))),
	}
}
