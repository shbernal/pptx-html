/**
 * The coverage report: what the pipeline can currently hold onto, counted.
 *
 * This is the project's progress metric, and it is deliberately not a pass/fail.
 * "How many round-trips are green" says nothing while the supported subset is
 * small — a converter that models three constructs and warns about everything
 * else is green and useless. What moves is the shape of the loss: notes leaving
 * `dropped`, causes leaving `unread`, slides leaving the carried lane.
 *
 * Classification is upstream's `Disposition` × `Cause`, never a local synonym.
 * `Cause` is also the triage: `unread` and `unwritable` are bugs in a specific
 * subsystem of the writer, `unsupported` is a property of OOXML or of the output
 * tier and will not be fixed by more converter work.
 */

import type { Cause, DeckIr, Disposition, FidelityNote } from 'pptx-ts/script'
import { isKnownNoteConstruct } from 'pptx-ts/script'

const DISPOSITIONS: Disposition[] = ['dropped', 'flattened', 'approximated']
const CAUSES: Cause[] = ['unread', 'unwritable', 'unsupported']

export interface CoverageRow {
	deck: string
	tier: string
	slides: number
	authored: number
	carried: number
	calls: number
	/** Note counts keyed `disposition/cause`; only non-zero cells are present. */
	cells: Record<string, number>
	/** Distinct constructs this deck's notes named, sorted. */
	constructs: string[]
}

export function coverageRow(deck: string, tier: string, ir: DeckIr): CoverageRow {
	const cells: Record<string, number> = {}
	for (const note of ir.fidelity) {
		const key = `${note.disposition}/${note.cause}`
		cells[key] = (cells[key] ?? 0) + 1
	}
	return {
		deck,
		tier,
		slides: ir.slides.length,
		authored: ir.slides.filter((slide) => slide.source === 'authored').length,
		carried: ir.slides.filter((slide) => slide.source === 'carried').length,
		calls: ir.slides.reduce((total, slide) => total + slide.calls.length, 0),
		cells,
		constructs: [...new Set(ir.fidelity.map((note) => note.construct))].toSorted(),
	}
}

/**
 * Constructs named by a note that the differ has no matcher for.
 *
 * A note whose construct the matcher table does not know can never declare
 * anything: it sits in `unmatchedNotes` forever while the difference it was
 * meant to excuse is reported as a defect. Cross-checking against upstream's own
 * table is what stops that from hiding.
 *
 * `isKnownNoteConstruct` rather than membership of `knownNoteConstructs()`: a
 * note recorded against a layout's shape carries the `layout.` prefix and is
 * deliberately absent from that list, so a set lookup calls every one of them
 * unmatchable. The predicate is the half of the pair that resolves the prefix.
 */
export function unmatchableConstructs(notes: FidelityNote[]): string[] {
	return [...new Set(notes.map((note) => note.construct))]
		.filter((construct) => !isKnownNoteConstruct(construct))
		.toSorted()
}

function pad(value: string, width: number): string {
	return value.length >= width ? value : value + ' '.repeat(width - value.length)
}

/**
 * A stable text rendering, for snapshotting. Column order is fixed and rows keep
 * corpus order, so a diff on this file reads as a change in *support*.
 */
export function formatCoverage(rows: CoverageRow[]): string {
	const lines: string[] = []
	const nameWidth = Math.max(4, ...rows.map((row) => row.deck.length))

	lines.push('# deck coverage — slides (authored/carried), calls')
	for (const row of rows) {
		lines.push(
			`${pad(row.deck, nameWidth)}  ${pad(row.tier, 9)}  slides=${row.slides} (${row.authored}a/${row.carried}c)  calls=${row.calls}`
		)
	}

	lines.push('')
	lines.push('# fidelity notes by disposition × cause (whole corpus)')
	const totals: Record<string, number> = {}
	for (const row of rows) {
		for (const [key, count] of Object.entries(row.cells)) totals[key] = (totals[key] ?? 0) + count
	}
	lines.push(`${pad('', 13)}${CAUSES.map((cause) => pad(cause, 13)).join('')}`)
	for (const disposition of DISPOSITIONS) {
		const cells = CAUSES.map((cause) => pad(String(totals[`${disposition}/${cause}`] ?? 0), 13)).join('')
		lines.push(`${pad(disposition, 13)}${cells}`)
	}

	lines.push('')
	lines.push('# constructs named, by deck')
	for (const row of rows) {
		lines.push(`${pad(row.deck, nameWidth)}  ${row.constructs.join(', ') || '—'}`)
	}

	return lines.join('\n')
}
