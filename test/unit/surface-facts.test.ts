import { ZipWriter } from 'pptx-ts/zip'
import { describe, expect, it } from 'vitest'
import { packageFacts } from '../surface/facts'

function zip(parts: Record<string, string | Uint8Array>): Uint8Array {
	const writer = new ZipWriter()
	for (const [path, data] of Object.entries(parts)) writer.add(path, data)
	return writer.toBytes()
}

/** A deck holding one embedded workbook, whose core properties carry `modified`. */
function deckWithWorkbook(modified: string, value = '3'): Uint8Array {
	const workbook = zip({
		'docProps/core.xml': `<cp:coreProperties><dcterms:modified>${modified}</dcterms:modified></cp:coreProperties>`,
		'xl/worksheets/sheet1.xml': `<worksheet><sheetData><row><c><v>${value}</v></c></row></sheetData></worksheet>`,
	})
	return zip({ 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx': workbook })
}

describe('packageFacts', () => {
	it("ignores an embedded package's docProps, as it does the deck's", async () => {
		const early = await packageFacts(deckWithWorkbook('2026-10-12T10:00:00Z'))
		const late = await packageFacts(deckWithWorkbook('2026-10-12T10:00:01Z'))
		expect([...late]).toEqual([...early])
	})

	it("reads an embedded package's content as facts under its own part kind", async () => {
		const facts = await packageFacts(deckWithWorkbook('2026-10-12T10:00:00Z', '5'))
		expect(facts).toContain(
			'ppt/embeddings/Microsoft_Excel_Worksheet#.xlsx!xl/worksheets/sheet#.xml⟩worksheet/sheetData/row/c/v#=5'
		)
	})
})
