/**
 * Minimal structural types for the slice of `@shbernal/ts-pptx` that the emit
 * layer drives. The emit layer is pure and isomorphic, so it depends on these
 * shapes rather than on the full writer types — this keeps the IR→ts-pptx
 * boundary decoupled and unit-testable against an injected/mock writer.
 *
 * Options objects (`addImage`/`addShape`/`addTable`/`addText`) are passed through
 * to ts-pptx opaquely, so they are typed as `Record<string, unknown>` rather
 * than mirroring its option unions.
 */

/** A ts-pptx writer instance, narrowed to what emit reads off it. */
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
