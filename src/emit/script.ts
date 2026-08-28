/**
 * `DeckIr` → `.pptx`. The fourth leg of the loop.
 *
 * This is an **interpreter for the tier `printScript` prints**, not a second
 * design: upstream's printer emits `fromTemplate` + one generator per layout run
 * + `appendSlides`, and this makes the same calls directly. That equivalence is
 * what makes `scriptTierNotes` the honest set of losses to declare for anything
 * that goes through here — a lane that declared a *different* set would both
 * excuse defects it really has and report differences it legitimately declared.
 *
 * It speaks `DeckIr` and nothing else. `RenderIr` is the paint model and must not
 * reach the writer: a renderer that can influence what gets emitted is a renderer
 * that can lose something silently, and the two models are kept apart precisely
 * so that cannot happen. Edits made in the rendered document reach this function
 * as a *patched `DeckIr`* — see `parse/edits.ts` — never as geometry.
 *
 * ## Emit is a function of the IR
 *
 * Nothing here measures text, reads a font metric or consults the host. The same
 * `DeckIr` and the same template produce the same deck on any machine, which is
 * what makes `IR₂ ≡ IR₁ ⇒ pptx₂ ≡ pptx₁` — and therefore Invariant R — mean
 * anything at all. "Equal" is normalized-model equality: `canonicalDeckIr`
 * absorbs `rId` numbering, `cNvPr` ids and zip ordering, so those vary legally
 * and are not defects to chase.
 */

import TsPptx from '@shbernal/ts-pptx'
import { type LayoutHandle, Presentation } from '@shbernal/ts-pptx/read'
import { type AssetIr, type CallIr, type DeckIr, type IrValue, isAssetRef, type SlideIr } from '@shbernal/ts-pptx/script'
import { EMU_PER_IN } from '../constants'

/**
 * Replay a `DeckIr` against a template package.
 *
 * `source` is the package the IR was read from, loaded separately from
 * `template`: `importSlide` copies *out of* a live `Presentation`, and
 * `fromTemplate` has already stripped the slides from the destination's copy of
 * the same bytes. It is optional only because a deck with no carried slide never
 * reaches for it — one that does and has none is a failure, not a fallback.
 */
export async function emitDeckIr(ir: DeckIr, template: Uint8Array, source?: Presentation): Promise<Uint8Array> {
	const destination = await Presentation.fromTemplate(template)
	const assets = new Map(ir.assets.map((asset) => [asset.name, asset]))

	// A batch is a run of consecutive slides sharing a layout: `appendSlides` binds
	// one layout per call, and a carried slide is not authored at all, so either
	// change ends the run. Getting this wrong reorders the deck.
	let batch: SlideIr[] = []
	const flush = async (): Promise<void> => {
		const pending = batch
		// Destructured rather than length-checked: `first` is the slide the layout
		// binds to, and taking it here is what says an empty batch has nothing to
		// bind, in a form the compiler can follow.
		const [first] = pending
		if (first === undefined) return
		batch = []
		const generator = generatorFor(ir)
		for (const slide of pending) authorSlide(generator, slide, assets)
		await destination.appendSlides(generator, { layout: layoutBinding(destination, first) })
	}

	for (const slide of ir.slides) {
		if (slide.source === 'carried') {
			await flush()
			if (source === undefined) {
				throw new Error(
					`slide ${slide.number} is carried, so emitting it needs the source package; pass one rather than transcribing a slide the write API cannot express`
				)
			}
			destination.importSlide(source, slide.number - 1, { importNotes: true })
			continue
		}
		if (batch[0] && batch[0].layout?.index !== slide.layout?.index) await flush()
		batch.push(slide)
	}
	await flush()

	// `save()`, not `toBytes()`. The 3.3.0 rename replaced `stream()` on the
	// *writer*; a `Presentation` from the read/edit side has always spelled it
	// this way, and the two are different objects.
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
		width: ir.slideSize.widthEmu / EMU_PER_IN,
		height: ir.slideSize.heightEmu / EMU_PER_IN,
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
 * bytes are what the IR carries, and a path would bind this to a file layout it
 * knows nothing about.
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
	const CHUNK = 0x8000
	let binary = ''
	for (let offset = 0; offset < bytes.length; offset += CHUNK) {
		binary += String.fromCharCode(...bytes.subarray(offset, offset + CHUNK))
	}
	return btoa(binary)
}
