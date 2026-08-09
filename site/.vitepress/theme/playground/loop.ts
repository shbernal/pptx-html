/**
 * The four legs, driven from a browser, with the bookkeeping the page shows.
 *
 * Framework-free on purpose: everything here is `pptx-html`'s public API used the
 * way a consumer would use it, so it is the part of the playground worth
 * typechecking. The Vue side owns pixels and nothing else.
 *
 * Two rules this module exists to keep:
 *
 * - **A stage never reports a state it did not earn.** A leg that throws is
 *   recorded as failed, and every leg after it is marked as not having run rather
 *   than left looking idle.
 * - **The source bytes are held for the whole session.** Masters, layouts, theme
 *   and every carried slide's XML live in the original package and nowhere else,
 *   so `emitDeck` needs them however many times the visitor edits.
 */

import { emitDeck, IslandError, importDeck, type ParsedDeck, parseDeck, renderDeck } from 'pptx-html'

export type StageName = 'importDeck' | 'renderDeck' | 'parseDeck' | 'emitDeck'
export type StageStatus = 'idle' | 'running' | 'ok' | 'failed' | 'blocked'

export interface Stage {
	name: StageName
	status: StageStatus
	/** Wall-clock for this leg, measured across the call and nothing else. */
	ms: number | null
	/** One line about what the leg produced. Empty until it has produced it. */
	note: string
	error: string | null
}

export const STAGE_NAMES: readonly StageName[] = ['importDeck', 'renderDeck', 'parseDeck', 'emitDeck']

export function idleStages(): Stage[] {
	return STAGE_NAMES.map((name) => ({ name, status: 'idle', ms: null, note: '', error: null }))
}

/**
 * Put a rendered document on screen and hand back the live `Document` it became.
 *
 * Supplied by the component, because the loop does not own the page. It matters
 * that this returns the *live* document rather than a re-parse of the string: the
 * document is what the visitor types into, and reading it back is the whole
 * demonstration.
 */
export type Present = (html: string) => Promise<Document>

export interface DeckSession {
	/** What to call the deck in the UI and in the download filename. */
	label: string
	source: Uint8Array
	html: string
	doc: Document
	/** What the renderer could not draw faithfully. Reported, never swallowed. */
	renderWarnings: string[]
	parsed: ParsedDeck
	emitted: Uint8Array
}

type Presented = Pick<DeckSession, 'label' | 'source' | 'html' | 'doc' | 'renderWarnings'>

/** Run all four legs over a deck's bytes. */
export async function openDeck(
	label: string,
	source: Uint8Array,
	present: Present,
	stages: Stage[]
): Promise<DeckSession> {
	reset(stages)

	const imported = await runStage(
		stages,
		'importDeck',
		() => importDeck(source),
		(result) => `${result.render.slides.length} slide${result.render.slides.length === 1 ? '' : 's'}`
	)

	const rendered = await runStage(
		stages,
		'renderDeck',
		async () => {
			const output = await renderDeck(imported.render, {
				bytes: (asset) => imported.assets.bytesFor({ $asset: asset }),
			})
			return { html: output.html, renderWarnings: output.warnings, doc: await present(output.html) }
		},
		(result) => `${kb(result.html.length)} of HTML`
	)

	return await readBack({ label, source, ...rendered }, stages)
}

/**
 * Re-run the return path against the document as it now stands.
 *
 * Called after every edit. The first two legs are deliberately left alone — they
 * did run, their numbers are still true, and re-running them would erase the
 * visitor's edit by re-rendering over it.
 */
export async function reread(session: DeckSession, stages: Stage[]): Promise<DeckSession> {
	for (const stage of stages) {
		if (stage.name === 'parseDeck' || stage.name === 'emitDeck') {
			Object.assign(stage, { status: 'idle', ms: null, note: '', error: null })
		}
	}
	return await readBack(session, stages)
}

async function readBack(base: Presented, stages: Stage[]): Promise<DeckSession> {
	const parsed = await runStage(
		stages,
		'parseDeck',
		async () => {
			// The live document, not a re-parse of the string: `parseHtml` is handed a
			// function so the caller decides what "the document" means, and here it
			// means the one the visitor has been typing into.
			const result = await parseDeck(base.html, { parseHtml: () => base.doc })
			if (result.ir === null) throw new Error('the rendered document carried no island')
			return result
		},
		(result) => laneNote(result)
	)

	const emitted = await runStage(
		stages,
		'emitDeck',
		// `source` is not optional and not a nicety: it is where the masters,
		// layouts, theme and any carried slide's untouched XML come from.
		() => emitDeck(parsed, { source: base.source }),
		(result) => `${kb(result.bytes.length)} of .pptx`
	)

	return { ...base, parsed, emitted: emitted.bytes }
}

function laneNote(parsed: ParsedDeck): string {
	const lanes = new Map<string, number>()
	for (const slide of parsed.slides) lanes.set(slide.lane, (lanes.get(slide.lane) ?? 0) + 1)
	const edits = parsed.slides.reduce((total, slide) => total + slide.edits, 0)
	const summary = [...lanes].map(([lane, count]) => `${count}× ${lane}`).join(', ')
	return edits === 0 ? summary : `${summary} · ${edits} edit${edits === 1 ? '' : 's'}`
}

// ---------------------------------------------------------------------------
// Stage bookkeeping
// ---------------------------------------------------------------------------

async function runStage<T>(
	stages: Stage[],
	name: StageName,
	work: () => Promise<T>,
	summarize: (result: T) => string
): Promise<T> {
	const stage = stageNamed(stages, name)
	stage.status = 'running'
	const started = performance.now()
	try {
		const result = await work()
		stage.ms = performance.now() - started
		stage.note = summarize(result)
		stage.status = 'ok'
		return result
	} catch (error) {
		stage.ms = performance.now() - started
		stage.error = describe(error)
		stage.status = 'failed'
		// Everything downstream is not merely idle — it did not run, and saying so
		// is the difference between a failure and a page that looks half-finished.
		for (const later of stages.slice(STAGE_NAMES.indexOf(name) + 1)) {
			Object.assign(later, { status: 'blocked', ms: null, note: 'did not run', error: null })
		}
		throw error
	}
}

function stageNamed(stages: Stage[], name: StageName): Stage {
	const stage = stages.find((entry) => entry.name === name)
	if (!stage) throw new Error(`no stage called ${name}`)
	return stage
}

function reset(stages: Stage[]): void {
	for (const stage of stages) Object.assign(stage, { status: 'idle', ms: null, note: '', error: null })
}

/**
 * What to show a visitor when a leg throws.
 *
 * `IslandError` is named rather than flattened into its message, because its
 * three faults call for three different responses: `version` means re-render from
 * the source, `tampered` means the document was altered after it was written, and
 * `malformed` means this was never one of our documents.
 *
 * No other error's `name` is shown. Errors thrown from inside a minified
 * dependency arrive named things like `De`, and prefixing a real message with two
 * letters of mangled class name makes the failure look like a bug in the page
 * rather than the thing it actually says.
 */
export function describe(error: unknown): string {
	if (error instanceof IslandError) return `IslandError (${error.fault}) — ${error.message}`
	if (error instanceof Error) return error.message
	return String(error)
}

export function kb(bytes: number): string {
	return `${(bytes / 1024).toFixed(1)} KB`
}
