/**
 * Minimal structural types for the slice of `pptx-ts` that the
 * heuristic lane drives. The lane's emitter is pure and isomorphic, so it
 * depends on these shapes rather than on the full writer types — which keeps the
 * model → ts-pptx boundary decoupled and unit-testable against an injected mock,
 * and is also what lets `opts.pptxFactory` hand back something that is not a
 * `TsPptx` at all.
 *
 * Options objects (`addImage`/`addShape`/`addTable`/`addText`) are passed through
 * to ts-pptx opaquely, so they are typed as `object` rather than mirroring its
 * option unions — this lane never reads one back, it only builds and forwards it.
 *
 * Two details make a real ts-pptx `Slide` satisfy this type, and both are load
 * bearing. The add-* members are written with **method syntax**, which checks
 * parameters bivariantly; written as function-typed properties, a concrete
 * `addImage(options: ImageProps)` would be rejected under `strictFunctionTypes`
 * for accepting less than this type promises. And the parameter is `object`
 * rather than `Record<string, unknown>`, because ts-pptx's option types are
 * interfaces, and an interface has no implicit index signature — so it is not
 * assignable to a `Record`, in either direction. Widening the other way, by
 * mirroring ts-pptx's option unions here, is what this file exists to avoid.
 */

/**
 * A ts-pptx writer instance, narrowed to what the emitter reads off it: the four
 * members of the shape-type enum this lane actually emits.
 *
 * Named rather than left as `Record<string, string>`, which said "any key" and
 * meant "these four". A mock that supplies none of them now fails to typecheck
 * instead of reaching `addShape(undefined, …)` at run time.
 */
export interface PptxWriter {
	ShapeType: {
		custGeom: string
		line: string
		rect: string
		roundRect: string
	}
}

/**
 * A ts-pptx slide, narrowed to the add-* methods and background setter emit uses.
 *
 * `background` is optional because the concrete `Slide` declares it that way — a
 * slide that states no background has none. Requiring it here would mean a real
 * `Slide` did not satisfy this type, which defeats the point of a subset.
 *
 * It carries `| undefined` as well as the `?`, and the two are not the same thing
 * under `exactOptionalPropertyTypes`: the `?` alone means "present with this type,
 * or absent", which a `Slide` whose own `background` is typed `… | undefined`
 * does not satisfy. Widening here rather than dropping the compiler option, which
 * is what keeps the rest of this file honest about what may be missing.
 *
 * The `| string` is the writer's `BackgroundOption`, which is `BackgroundProps` or
 * a bare `Color`. This lane never sets that form, but a subset type has to admit
 * everything the concrete one does or the concrete one stops satisfying it.
 */
export interface PptxSlide {
	background?: { color?: string; path?: string; data?: string } | string | undefined
	addImage(opts: object): unknown
	addShape(type: string, opts: object): unknown
	addTable(rows: unknown, opts: object): unknown
	addText(text: unknown, opts: object): unknown
	addNotes?(notes: string): unknown
}

/**
 * A ts-pptx deck, narrowed to what the orchestrator sets and calls. Deliberately
 * the same treatment as `PptxSlide`: `opts.pptxFactory` may return a mock, so the
 * orchestrator must not depend on anything beyond this.
 */
export interface PptxDeck {
	layout: string
	author?: string
	subject?: string
	company?: string
	// No `lang`. The writer has no deck-level one, so declaring it here would let a
	// caller — and did let this lane — write a property that goes nowhere. The
	// language travels as a run option; see `slide.ts`.
	theme?: Record<string, unknown>
	/** Slide dimensions in EMU, read back after `layout` is set. */
	presLayout?: { width: number; height: number }
	addSlide(): PptxSlide
	/**
	 * The deck as bytes. Present as of ts-pptx 3.3.0, and optional so a
	 * `pptxFactory` mock predating it still satisfies this type. `deliverDeck`
	 * prefers it and falls back to `write` when it is absent: the fallback is
	 * deliberate compatibility, not legacy code awaiting removal.
	 */
	toBytes?(opts?: object): Promise<Uint8Array>
	write(opts: { outputType: 'base64' }): Promise<string>
}
