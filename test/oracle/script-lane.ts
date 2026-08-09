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
import { type LayoutHandle, Presentation } from '@shbernal/ts-pptx/read'
import {
	type AssetIr,
	type CallIr,
	type DeckIr,
	type IrValue,
	isAssetRef,
	type SlideIr,
} from '@shbernal/ts-pptx/script'
import { EMU_PER_INCH } from '../../src/ir/render'
import { type Loop, scriptTierNotes } from './roundtrip'

/**
 * The lane declares the tier's own losses and nothing else.
 *
 * It used to add one of its own — a table whose rows are **all** auto-height came
 * back with explicit heights, and `readModelToIr` excluded that case from its
 * `table.rowAuto` note explicitly. ts-pptx 3.0.0 fixed it
 * (https://github.com/shbernal/ts-pptx/issues/5), so `scriptTierNotes` now covers
 * it and a local compensation would be a second note for one loss.
 */
export const scriptLoop: Loop = async (input) => {
	const bytes = await emitDeckIr(input.ir, input.bytes, input.pres)
	return { bytes, notes: scriptTierNotes(input.ir) }
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
