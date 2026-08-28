/**
 * The pixel/point constants and the default font mapping.
 *
 * NOTE: the browser extractor (`heuristic/extractor.ts`) redefines PX_PER_IN /
 * PT_PER_PX internally on purpose — it is stringified and `eval`'d in another
 * realm, so it cannot import these. Keeping local constants with the same names
 * is what makes that comment checkable against something.
 *
 * ## Where the EMU constants are, and why they are not here
 *
 * `EMU_PER_IN` used to live here as a third name for a number `src/ir/render.ts`
 * already exported as `EMU_PER_INCH`. One number, two names, two files, both
 * live. The rule now:
 *
 * - **`EMU_PER_INCH` and `EMU_PER_POINT` are the model's own**, exported from
 *   `src/ir/render.ts` beside `inchesOf`/`emuOf`. `docs/architecture.md` makes EMU
 *   the model's stated unit, and a consumer reading `EMU_PER_INCH` off this
 *   package should not have its value depend on a transitive dependency's release.
 * - **Internal-only conversions take upstream's.** `src/emit/script.ts` may not
 *   import `src/ir/render.ts` at all — that is the emit boundary — so it takes
 *   `EMU_PER_INCH` from `@shbernal/ts-pptx`, which is the library that owns the
 *   format. `src/render/geometry.ts` does the same for `ANGLE_UNITS_PER_DEGREE`.
 *
 * `PX_PER_IN` and `PT_PER_PX` stay here rather than becoming upstream's
 * `DEFAULT_PX_PER_INCH` and `POINTS_PER_INCH / DEFAULT_PX_PER_INCH`, for the
 * extractor reason above.
 */

export const PX_PER_IN = 96
export const PT_PER_PX = 0.75

export const DEFAULT_FONT = {
	latin: 'Carlito',
	cjk: 'Microsoft YaHei',
	symbol: 'Segoe UI Symbol',
	emphasis: 'Liberation Sans Narrow',
}
