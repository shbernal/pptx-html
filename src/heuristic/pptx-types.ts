/**
 * Minimal structural types for the slice of `@shbernal/ts-pptx` that the
 * heuristic lane drives. The lane's emitter is pure and isomorphic, so it
 * depends on these shapes rather than on the full writer types — which keeps the
 * model → ts-pptx boundary decoupled and unit-testable against an injected mock,
 * and is also what lets `opts.pptxFactory` hand back something that is not a
 * `TsPptx` at all.
 *
 * Options objects (`addImage`/`addShape`/`addTable`/`addText`) are passed through
 * to ts-pptx opaquely, so they are typed as `Record<string, unknown>` rather
 * than mirroring its option unions.
 */

/** A ts-pptx writer instance, narrowed to what the emitter reads off it. */
export interface PptxWriter {
	/** Shape-type enum: `ShapeType.rect`, `.roundRect`, `.line`, `.custGeom`, … */
	ShapeType: Record<string, string>
}

/** A ts-pptx slide, narrowed to the add-* methods and background setter emit uses. */
export interface PptxSlide {
	background: { color?: string; path?: string; data?: string }
	addImage: (opts: Record<string, unknown>) => unknown
	addShape: (type: string, opts: Record<string, unknown>) => unknown
	addTable: (rows: unknown, opts: Record<string, unknown>) => unknown
	addText: (text: unknown, opts: Record<string, unknown>) => unknown
	addNotes?: (notes: string) => unknown
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
