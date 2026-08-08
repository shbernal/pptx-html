/**
 * The `import → emit` lane: half the loop, testable before any HTML exists.
 *
 * `readModelToIr` reads a deck into `DeckIr`; this replays that IR through the
 * write API and saves. Nothing renders, nothing parses, and both legs are
 * upstream code — which is the point. A difference here belongs to the
 * import/emit pair and cannot be misattributed to a renderer that has not been
 * written yet, and because neither leg is local, a failure is an upstream issue
 * to file rather than something to patch here.
 *
 * It is an **interpreter for the same tier `printScript` prints**, not a second
 * design. `printScript` emits `fromTemplate` + one generator per layout run +
 * `appendSlides`, and this makes the same calls directly, so `scriptTierNotes`
 * is the honest set of losses to declare (see `LoopOutput.notes`, where getting
 * this wrong both over- and under-excludes). Printing source and then evaluating
 * it would test a TypeScript emitter; this tests the mapping.
 *
 * It lives in `test/` rather than `src/` because emit is part 06's, and a lane
 * the oracle drives is not yet the package's emit path.
 */

import TsPptx from '@shbernal/ts-pptx'
import { isGraphicFrame, type LayoutHandle, Presentation } from '@shbernal/ts-pptx/read'
import {
	type AssetIr,
	type CallIr,
	type DeckIr,
	type FidelityNote,
	type IrValue,
	isAssetRef,
	type SlideIr,
} from '@shbernal/ts-pptx/script'
import { EMU_PER_INCH } from '../../src/ir/render'
import { type Loop, scriptTierNotes } from './roundtrip'

export const scriptLoop: Loop = async (input) => {
	const bytes = await emitDeckIr(input.ir, input.bytes, input.pres)
	return { bytes, notes: [...scriptTierNotes(input.ir), ...autoRowNotes(input.pres)] }
}

/**
 * The one loss this lane declares beyond the tier's own: a table whose rows are
 * **all** auto-height comes back with explicit heights.
 *
 * `a:tr/@h="0"` means "size to content", and `readModelToIr` correctly emits no
 * `rowH` for it — but `addTable` then divides the frame height evenly, so the
 * re-read reports three pinned rows where the source had three auto ones. The
 * table still looks identical; the *implicitness* is what is lost.
 *
 * Upstream already owns this construct and already notes the mixed case (some
 * rows auto, some not). It does not note the all-auto case, which is the more
 * common one, so that gap is a part-07 ask rather than something to fix here.
 * The note is scoped to the tables it actually applies to — a blanket one would
 * excuse a genuinely wrong `rowH` on any table in the deck.
 */
function autoRowNotes(pres: Presentation): FidelityNote[] {
	const notes: FidelityNote[] = []
	for (const slide of pres.slides) {
		for (const shape of slide.shapes) {
			if (!isGraphicFrame(shape)) continue
			const table = shape.table
			if (table === null || table.rows.length === 0) continue
			if (!table.rows.every((row) => (row.heightEmu ?? 0) === 0)) continue
			notes.push({
				slideNumber: slide.index + 1,
				shapeName: shape.name,
				construct: 'table.rowAuto',
				disposition: 'approximated',
				cause: 'unsupported',
				detail:
					'every row of this table is auto-height (a:tr/@h of 0), so the IR carries no rowH; addTable then divides the frame height evenly and the rows come back pinned to that share',
			})
		}
	}
	return notes
}

/**
 * Replay a `DeckIr` against the deck it was read from, used as a template.
 *
 * `source` is the same package as `template`, loaded separately, because
 * `importSlide` copies *out of* a live `Presentation` and `fromTemplate` has
 * already stripped the slides from the destination's copy.
 */
export async function emitDeckIr(ir: DeckIr, template: Uint8Array, source: Presentation): Promise<Uint8Array> {
	const destination = await Presentation.fromTemplate(template)
	const assets = new Map(ir.assets.map((asset) => [asset.name, asset]))

	// A batch is a run of consecutive slides sharing a layout: `appendSlides` binds
	// one layout per call, and a carried slide is not authored at all, so either
	// change ends the run. Getting this wrong reorders the deck.
	let batch: SlideIr[] = []
	const flush = async (): Promise<void> => {
		if (batch.length === 0) return
		const pending = batch
		batch = []
		const generator = generatorFor(ir)
		for (const slide of pending) authorSlide(generator, slide, assets)
		await destination.appendSlides(generator, { layout: layoutBinding(destination, pending[0]) })
	}

	for (const slide of ir.slides) {
		if (slide.source === 'carried') {
			await flush()
			destination.importSlide(source, slide.number - 1, { importNotes: true })
			continue
		}
		if (batch[0] && batch[0].layout?.index !== slide.layout?.index) await flush()
		batch.push(slide)
	}
	await flush()

	return destination.save()
}

/**
 * A generator sized to the template. `appendSlides` compares slide sizes exactly
 * and throws when they differ, so this has to round-trip the source EMU rather
 * than approximate it.
 */
function generatorFor(ir: DeckIr): TsPptx {
	const pptx = new TsPptx()
	pptx.defineLayout({
		name: 'source',
		width: ir.slideSize.widthEmu / EMU_PER_INCH,
		height: ir.slideSize.heightEmu / EMU_PER_INCH,
	})
	pptx.layout = 'source'
	return pptx
}

/**
 * Bind by name where the name is unambiguous, by gallery position where it is
 * not. `appendSlides` throws on an ambiguous name rather than choosing, and a
 * multi-master deck routinely repeats layout names — which is exactly what
 * `SlideLayoutIr.nameIsUnique` is there to say.
 */
function layoutBinding(destination: Presentation, first: SlideIr): string | LayoutHandle {
	const layout = first.layout
	const gallery = destination.layouts()
	if (!layout) {
		const fallback = gallery[0]
		if (!fallback) throw new Error('the template has no layouts to bind to')
		return fallback
	}
	if (layout.nameIsUnique) return layout.name
	const handle = gallery[layout.index]
	if (!handle) throw new Error(`the template has no layout at gallery position ${layout.index}`)
	return handle
}

/**
 * Slide-level properties are assignments on the write API, not calls; calls
 * follow in z-order.
 *
 * `CallIr.method` names a real method and `args` are its real arguments, so this
 * is a dynamic dispatch by design — the IR's whole claim is that a printer is a
 * formatter. The two untyped views are that claim taken at face value; the
 * `typeof` guard below is what turns a broken claim into a named failure rather
 * than a `TypeError`.
 */
function authorSlide(generator: TsPptx, slide: SlideIr, assets: ReadonlyMap<string, AssetIr>): void {
	const authored = generator.addSlide()
	const properties = authored as unknown as Record<string, unknown>
	const methods = authored as unknown as Record<string, unknown>

	if (slide.hidden) properties.hidden = true
	if (slide.background) properties.background = hydrate(slide.background as unknown as IrValue, assets)
	if (slide.transition) properties.transition = hydrate(slide.transition as unknown as IrValue, assets)
	if (slide.notesText !== undefined) invoke(methods, { method: 'addNotes', args: [slide.notesText] }, assets)
	for (const call of slide.calls) invoke(methods, call, assets)
}

function invoke(slide: Record<string, unknown>, call: CallIr, assets: ReadonlyMap<string, AssetIr>): void {
	const method = slide[call.method]
	if (typeof method !== 'function') throw new Error(`the write API has no method ${call.method}`)
	;(method as (...args: unknown[]) => unknown).call(slide, ...call.args.map((arg) => hydrate(arg, assets)))
}

/**
 * Turn every `AssetRef` in an argument tree back into bytes.
 *
 * A `data:` URI rather than a path, for the reason upstream inlines them: the
 * bytes are what the IR carries, and a path would bind this lane to a file
 * layout it knows nothing about.
 */
function hydrate(value: IrValue, assets: ReadonlyMap<string, AssetIr>): unknown {
	if (isAssetRef(value)) {
		const asset = assets.get(value.$asset)
		if (!asset) throw new Error(`the IR references asset ${value.$asset}, which it does not carry`)
		return `data:${asset.contentType};base64,${base64(asset.bytes)}`
	}
	if (Array.isArray(value)) return value.map((entry) => hydrate(entry, assets))
	if (value !== null && typeof value === 'object') {
		return Object.fromEntries(Object.entries(value).map(([key, entry]) => [key, hydrate(entry, assets)]))
	}
	return value
}

/** Node and the browser disagree about `Buffer`; this needs neither. */
function base64(bytes: Uint8Array): string {
	let binary = ''
	for (const byte of bytes) binary += String.fromCharCode(byte)
	return btoa(binary)
}
