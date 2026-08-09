/**
 * The loop, end to end, in one file.
 *
 * The README's claim is that a deck can leave PPTX, be edited as a web page, and
 * come back a deck rather than a picture of one. This script is that claim made
 * executable: it writes a deck, runs all four legs over it, retypes one run
 * through the sanctioned surface, and writes the result back out.
 *
 * ```bash
 * pnpm run build   # the example imports the package by name, so dist/ must exist
 * pnpm run example
 * ```
 *
 * It leaves three files in `examples/out/`: the deck it started from, the HTML it
 * became, and the deck it came back as. Open the first and the last in PowerPoint
 * — the only difference is the word this script edited.
 *
 * Two things are worth watching in the output rather than in the files.
 *
 * The **lane** printed per slide is the return path's own account of how it read
 * the document: `exact` means the model came back as it was rendered, and the
 * other three (`reconciled`, `drifted`, `heuristic`) each mean something weaker.
 * A conversion that degrades quietly is the failure mode this project is arranged
 * against, so the lane is reported, not inferred from the fact that a file
 * appeared.
 *
 * The **source bytes** handed to `emitDeck` are not optional. Masters, layouts,
 * the theme, and the untouched XML of any carried slide live in the original
 * package and nowhere else — the document carries the edits, the caller supplies
 * the substance.
 */

import { mkdir, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { TsPptx } from '@shbernal/ts-pptx'
import { emitDeck, importDeck, parseDeck, renderDeck } from 'pptx-html'

const OUT = join(dirname(fileURLToPath(import.meta.url)), 'out')

/** A two-slide deck to start from. Any `.pptx` on disk would do just as well. */
async function sourceDeck() {
	const pptx = new TsPptx()
	pptx.layout = 'LAYOUT_16x9'
	pptx.title = 'pptx-html example'

	pptx.addSlide().addText([{ text: 'Hello', options: { bold: true, color: '250F6B' } }, { text: ' from PPTX' }], {
		x: 0.5,
		y: 0.5,
		w: 8,
		h: 1.5,
		fontSize: 36,
		objectName: 'title',
	})

	pptx.addSlide().addShape('roundRect', {
		x: 1,
		y: 1,
		w: 4,
		h: 2,
		fill: { color: 'DDE3F0' },
		line: { color: '250F6B', width: 2 },
		objectName: 'card',
	})

	return await pptx.write({ outputType: 'uint8array' })
}

await mkdir(OUT, { recursive: true })

// ── 1. import ───────────────────────────────────────────────────────────────
// Bytes in, paint model out. `assets` holds the deck's images, keyed so the
// renderer can ask for them without the model carrying binary itself.
const source = await sourceDeck()
await writeFile(join(OUT, 'source.pptx'), source)
const imported = await importDeck(source)
console.log(`imported ${imported.render.slides.length} slides`)

// ── 2. render ───────────────────────────────────────────────────────────────
// The model becomes a document with two channels: visible SVG, and the JSON
// island beside it. Only the island is trusted on the way back.
const { html } = await renderDeck(imported.render, {
	bytes: (asset) => imported.assets.bytesFor({ $asset: asset }),
})
await writeFile(join(OUT, 'deck.html'), html)
console.log(`rendered ${html.length} bytes of HTML`)

// ── 3. parse ────────────────────────────────────────────────────────────────
// `parseHtml: null` reads the island and skips the visual channel. That is the
// honest reading in Node: a string that has never been a live document cannot
// have been typed into, so there are no DOM edits to collect. A browser caller
// passes a real `ParentNode` here and gets the full four-lane behaviour.
const parsed = await parseDeck(html, { parseHtml: null })
if (parsed.ir === null) throw new Error('the rendered document carried no island')
for (const slide of parsed.slides) {
	console.log(`slide ${slide.number}: lane=${slide.lane} edits=${slide.edits}`)
}

// ── 4. edit, through the declared surface ───────────────────────────────────
// `project(ir)` names exactly this set; anything outside it is not an edit the
// loop will carry, and changing it here would be silently dropped rather than
// silently emitted.
const firstRun = parsed.ir.slides[0]?.nodes.find((node) => node.kind === 'shape' && node.text)?.text?.paragraphs[0]
	?.runs[0]
if (!firstRun) throw new Error('the example deck lost its first run')
console.log(`editing ${JSON.stringify(firstRun.text)} -> "Goodbye"`)
firstRun.text = 'Goodbye'

// ── 5. emit ─────────────────────────────────────────────────────────────────
// `source` is required: it is where the masters, layouts, theme and any carried
// slide's XML come from. Emitting without it would mean inventing them.
const emitted = await emitDeck(parsed, { source })
await writeFile(join(OUT, 'round-tripped.pptx'), emitted.bytes)
console.log(`wrote examples/out/round-tripped.pptx (${emitted.bytes.length} bytes)`)

// ── 6. prove it ─────────────────────────────────────────────────────────────
// An example that stops at "a file appeared" would pass just as happily if the
// edit had gone nowhere — which is exactly the failure this project treats as
// the dangerous one. So read the emitted deck back through the same first leg
// and look at the run that was changed.
const back = await importDeck(emitted.bytes)
const backRun = back.render.slides[0]?.nodes.find((node) => node.kind === 'shape' && node.text)?.text?.paragraphs[0]
	?.runs[0]
console.log(`re-imported first run: ${JSON.stringify(backRun?.text)}`)

// The parse warning below is worth reading rather than ignoring. It says the
// *document's* channel was not read — true, and the reason `parseHtml: null` is
// passed above. It does not describe the edit this script made, which was made
// to the model object and travelled through `emitDeck` on the island's terms.
if (parsed.warnings.length > 0) console.log('warnings:', parsed.warnings)
