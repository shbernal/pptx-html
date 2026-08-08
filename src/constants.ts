/**
 * Shared unit constants and the default font mapping. Used by the emit layer and
 * the deck orchestrator.
 *
 * NOTE: the browser extractor (`extract/extractor.ts`) redefines PX_PER_IN /
 * PT_PER_PX internally on purpose — it is stringified and `eval`'d in another
 * realm, so it cannot import these.
 */

export const PX_PER_IN = 96
export const PT_PER_PX = 0.75
export const EMU_PER_IN = 914400

export const DEFAULT_FONT = {
	latin: 'Carlito',
	cjk: 'Microsoft YaHei',
	symbol: 'Segoe UI Symbol',
	emphasis: 'Liberation Sans Narrow',
}
