/**
 * Minimal structural types for the slice of `@shbernal/ts-pptx` that the
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
 */
export interface PptxSlide {
	background?: { color?: string; path?: string; data?: string }
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
	lang?: string
	theme?: Record<string, unknown>
	/** Slide dimensions in EMU, read back after `layout` is set. */
	presLayout?: { width: number; height: number }
	addSlide: () => PptxSlide
	write: (opts: { outputType: 'base64' }) => Promise<string>
}
