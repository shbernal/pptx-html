<script setup lang="ts">
import { onMounted, reactive, ref, shallowRef } from 'vue'
import { download } from '../download.ts'
import { type DeckSession, describe, idleStages, kb, openDeck, reread } from './loop.ts'
import PipelineStrip from './PipelineStrip.vue'
import { type Sample, sampleNamed, samples } from './samples.ts'
import SurfacePanel from './SurfacePanel.vue'
import { deleteNode, setProp, setText, type SlideRow, srgb, surfaceOf } from './surface.ts'

/** Straight from `reconcile.ts`, because paraphrasing a guarantee weakens it. */
const LANES: Record<string, string> = {
	exact: "the slide's surface came back identical; Invariant R holds outright",
	reconciled: 'a sanctioned edit; Invariant R holds for everything else',
	drifted: "something outside the surface changed; the island's value stands",
	heuristic: 'there was no island at all',
}

const stages = reactive(idleStages())
const sampleList = ref<Sample[]>([])
const session = shallowRef<DeckSession | null>(null)
const rows = shallowRef<SlideRow[]>([])
const active = ref<Sample | null>(null)
const activeLabel = ref<string | null>(null)
const busy = ref(false)
const failure = ref<string | null>(null)
const dragging = ref(false)
const frame = ref<HTMLIFrameElement | null>(null)

onMounted(() => {
	try {
		sampleList.value = samples()
	} catch (error) {
		failure.value = describe(error)
		return
	}
	// `?deck=` is how the fidelity ledger links a row to the deck it measured. It
	// resolves against the whole corpus, not just the curated picker — but an
	// unknown name says so rather than quietly loading something else.
	const wanted = new URLSearchParams(window.location.search).get('deck')
	if (wanted === null) return
	const deck = sampleNamed(wanted)
	if (deck === null) {
		failure.value = `There is no deck called "${wanted}" in the corpus, so nothing was loaded. Pick one above.`
		return
	}
	if (!sampleList.value.some((sample) => sample.name === deck.name)) sampleList.value = [...sampleList.value, deck]
	void pick(deck)
})

/**
 * Put the rendered document in the preview and hand back the live document.
 *
 * No `sandbox` attribute, and that is a decision rather than an omission. The
 * loop needs `contentDocument` (same-origin) and the document's own hydration
 * script (scripts allowed), which together are what a sandbox would have to grant
 * back — `sandbox="allow-scripts allow-same-origin"` is a label, not a boundary.
 * What the frame is really for is style isolation: a slide restyled by VitePress's
 * stylesheet would be a fake preview. The document is written by this project's
 * renderer, which escapes every string it takes from a deck, so a dropped file's
 * content is data here and never markup.
 */
async function present(html: string): Promise<Document> {
	const element = frame.value
	if (element === null) throw new Error('the preview frame is not on the page')
	await new Promise<void>((resolve) => {
		element.addEventListener('load', () => resolve(), { once: true })
		element.srcdoc = html
	})
	const doc = element.contentDocument
	if (doc === null) throw new Error('the preview frame produced no document')
	// Typing into a slide is an edit like any other, so it re-runs the return path
	// on its own. Debounced because each pass is a full `parseDeck` + `emitDeck`,
	// and a keystroke is not a reason to rewrite a deck.
	doc.addEventListener('input', () => {
		clearTimeout(typingTimer)
		typingTimer = setTimeout(() => void edit(() => {}), 600)
	})
	return doc
}

let typingTimer: ReturnType<typeof setTimeout> | undefined

async function open(label: string, bytes: Uint8Array): Promise<void> {
	busy.value = true
	failure.value = null
	session.value = null
	rows.value = []
	activeLabel.value = label
	try {
		const opened = await openDeck(label, bytes, present, stages)
		session.value = opened
		rows.value = surfaceOf(irOf(opened))
	} catch (error) {
		failure.value = describe(error)
	} finally {
		busy.value = false
	}
}

function irOf(deck: DeckSession) {
	const ir = deck.parsed.ir
	if (ir === null) throw new Error('the parsed deck carried no model')
	return ir
}

async function pick(sample: Sample): Promise<void> {
	active.value = sample
	busy.value = true
	try {
		const bytes = await sample.build()
		await open(sample.name, bytes)
	} catch (error) {
		failure.value = describe(error)
		busy.value = false
	}
}

async function accept(file: File | undefined): Promise<void> {
	if (!file) return
	active.value = null
	await open(file.name.replace(/\.pptx$/i, ''), new Uint8Array(await file.arrayBuffer()))
}

/**
 * Every edit takes the same route: change the live document, then read it back.
 *
 * The panel is not a second way in. It writes the same spans a visitor types
 * into, and `parseDeck` is what decides what any of it meant.
 */
async function edit(mutate: (doc: Document) => void): Promise<void> {
	const current = session.value
	if (current === null || busy.value) return
	busy.value = true
	failure.value = null
	try {
		mutate(current.doc)
		const next = await reread(current, stages)
		session.value = next
		rows.value = surfaceOf(irOf(next))
	} catch (error) {
		failure.value = describe(error)
	} finally {
		busy.value = false
	}
}

function onDrop(event: DragEvent): void {
	dragging.value = false
	void accept(event.dataTransfer?.files?.[0])
}
</script>

<template>
	<div class="pxh-playground">
		<p class="pxh-privacy">
			Everything below runs in this tab. No deck is uploaded, and there is no server in the loop — the same four
			functions a consumer calls are running in your browser.
		</p>

		<!-- ── deck source ─────────────────────────────────────────────────── -->
		<section class="pxh-sources">
			<h3>1 · Pick a deck</h3>
			<p class="pxh-hint">
				These are decks from the round-trip oracle's own corpus, not demo decks written for this page. A
				<span class="pxh-tier is-hard">hard</span> deck is one the oracle watches rather than gates.
			</p>
			<ul class="pxh-samples">
				<li v-for="sample in sampleList" :key="sample.name">
					<button
						type="button"
						class="pxh-sample"
						:class="{ 'is-active': active?.name === sample.name }"
						:disabled="busy"
						@click="pick(sample)"
					>
						<span class="pxh-sample-head">
							<code>{{ sample.name }}</code>
							<span class="pxh-tier" :class="`is-${sample.tier}`">{{ sample.tier }}</span>
						</span>
						<span class="pxh-sample-why">{{ sample.why }}</span>
					</button>
				</li>
			</ul>

			<div
				class="pxh-drop"
				:class="{ 'is-dragging': dragging }"
				@dragover.prevent="dragging = true"
				@dragleave="dragging = false"
				@drop.prevent="onDrop"
			>
				<label>
					<input type="file" accept=".pptx" :disabled="busy" @change="accept(($event.target as HTMLInputElement).files?.[0])" />
					<span>…or drop a <code>.pptx</code> of your own here.</span>
				</label>
				<p>
					Decks authored in PowerPoint are the project's <strong>second tier</strong>: the guarantee is scoped to
					decks written by <code>@shbernal/ts-pptx</code>, and PowerPoint-authored files are not gated by the
					oracle. Anything imperfect you see is the documented state of the project, shown rather than hidden.
				</p>
			</div>
		</section>

		<!-- ── the loop ────────────────────────────────────────────────────── -->
		<section>
			<h3>2 · The loop</h3>
			<PipelineStrip :stages="stages" />
			<p v-if="failure" class="pxh-failure">{{ failure }}</p>
		</section>

		<!-- ── preview + surface ───────────────────────────────────────────── -->
		<section>
			<h3>3 · The document, and what may be edited in it</h3>
			<p class="pxh-hint">
				The preview is the real rendered document. Run text is <code>contenteditable</code>, so you can type
				straight into a slide — the panel beside it writes to the same spans. Both are read back by the same
				<code>parseDeck</code> call. Clearing a property returns the run to inheriting it.
			</p>

			<div class="pxh-split">
				<div class="pxh-preview">
					<iframe ref="frame" title="Rendered deck"></iframe>
					<p v-if="session === null" class="pxh-empty">Pick a deck to render one.</p>
				</div>

				<div class="pxh-side">
					<button type="button" class="pxh-reread" :disabled="session === null || busy" @click="edit(() => {})">
						Read the document back
					</button>
					<SurfacePanel
						:slides="rows"
						@text="(address, value) => edit((doc) => setText(doc, address, value))"
						@bold="(address, value) => edit((doc) => setProp(doc, address, 'bold', value ? true : undefined))"
						@italic="(address, value) => edit((doc) => setProp(doc, address, 'italic', value ? true : undefined))"
						@size="(address, value) => edit((doc) => setProp(doc, address, 'sizePt', value ?? undefined))"
						@color="(address, value) => edit((doc) => setProp(doc, address, 'color', value === null ? undefined : srgb(value)))"
						@remove="(id) => edit((doc) => deleteNode(doc, id))"
					/>
				</div>
			</div>
		</section>

		<!-- ── what came back ──────────────────────────────────────────────── -->
		<section v-if="session">
			<h3>4 · What came back</h3>

			<ul class="pxh-lanes">
				<li v-for="slide in session.parsed.slides" :key="slide.number">
					<span class="pxh-lane" :class="`is-${slide.lane}`" :title="LANES[slide.lane]">{{ slide.lane }}</span>
					<span class="pxh-lane-slide">slide {{ slide.number }}</span>
					<span v-if="slide.edits > 0" class="pxh-lane-edits">
						{{ slide.edits }} edit{{ slide.edits === 1 ? '' : 's' }}
					</span>
					<span v-for="note in slide.notes" :key="note" class="pxh-lane-note">{{ note }}</span>
				</li>
			</ul>

			<details v-if="session.renderWarnings.length > 0 || session.parsed.warnings.length > 0" class="pxh-warnings">
				<summary>
					{{ session.renderWarnings.length + session.parsed.warnings.length }} warning(s) — shown, not swallowed
				</summary>
				<ul>
					<li v-for="warning in session.renderWarnings" :key="warning"><code>renderDeck</code> — {{ warning }}</li>
					<li v-for="warning in session.parsed.warnings" :key="warning"><code>parseDeck</code> — {{ warning }}</li>
				</ul>
			</details>

			<div class="pxh-downloads">
				<button type="button" @click="download(session.source, `${session.label}.pptx`)">
					Download the original ({{ kb(session.source.length) }})
				</button>
				<button type="button" class="pxh-primary" @click="download(session.emitted, `${session.label}-round-tripped.pptx`)">
					Download the round-tripped deck ({{ kb(session.emitted.length) }})
				</button>
			</div>
			<p class="pxh-hint">
				Open both in PowerPoint. The only difference should be the words you changed — that comparison is the
				claim; one file on its own is just a file.
			</p>
		</section>
	</div>
</template>

<style scoped>
.pxh-playground h3 {
	margin: 32px 0 6px;
	font-size: 15px;
	border: none;
	padding: 0;
}

.pxh-privacy {
	margin: 0;
	padding: 10px 14px;
	border-left: 3px solid var(--vp-c-brand-1);
	background: var(--vp-c-bg-soft);
	border-radius: 0 8px 8px 0;
	font-size: 13.5px;
}

.pxh-hint {
	font-size: 13px;
	color: var(--vp-c-text-2);
	line-height: 1.6;
}

.pxh-samples {
	display: grid;
	grid-template-columns: repeat(auto-fill, minmax(230px, 1fr));
	gap: 8px;
	margin: 12px 0;
	padding: 0;
	list-style: none;
}

.pxh-sample {
	display: block;
	width: 100%;
	height: 100%;
	text-align: left;
	border: 1px solid var(--vp-c-divider);
	border-radius: 8px;
	padding: 8px 10px;
	background: var(--vp-c-bg-soft);
	transition: border-color 0.2s;
}

.pxh-sample:hover:not(:disabled) {
	border-color: var(--vp-c-brand-1);
}

.pxh-sample.is-active {
	border-color: var(--vp-c-brand-1);
	background: var(--vp-c-brand-soft);
}

.pxh-sample:disabled {
	opacity: 0.6;
}

.pxh-sample-head {
	display: flex;
	align-items: center;
	gap: 8px;
}

.pxh-sample-head code {
	background: none;
	padding: 0;
	font-size: 13px;
}

.pxh-sample-why {
	display: block;
	margin-top: 3px;
	font-size: 12px;
	line-height: 1.5;
	color: var(--vp-c-text-3);
}

.pxh-tier {
	font-size: 10.5px;
	text-transform: uppercase;
	letter-spacing: 0.06em;
	border-radius: 20px;
	padding: 1px 7px;
	border: 1px solid var(--vp-c-divider);
	color: var(--vp-c-text-3);
}

.pxh-tier.is-hard {
	border-color: var(--vp-c-warning-1);
	color: var(--vp-c-warning-1);
}

.pxh-drop {
	border: 1px dashed var(--vp-c-divider);
	border-radius: 8px;
	padding: 12px 14px;
}

.pxh-drop.is-dragging {
	border-color: var(--vp-c-brand-1);
	background: var(--vp-c-brand-soft);
}

.pxh-drop p {
	margin: 8px 0 0;
	font-size: 12.5px;
	line-height: 1.6;
	color: var(--vp-c-text-2);
}

.pxh-drop input {
	font-size: 13px;
}

.pxh-failure {
	margin-top: 10px;
	padding: 10px 14px;
	border-left: 3px solid var(--vp-c-danger-1);
	background: var(--vp-c-danger-soft);
	border-radius: 0 8px 8px 0;
	font-size: 13px;
	white-space: pre-wrap;
}

.pxh-split {
	display: grid;
	grid-template-columns: minmax(0, 3fr) minmax(0, 2fr);
	gap: 16px;
	align-items: start;
}

@media (max-width: 860px) {
	.pxh-split {
		grid-template-columns: minmax(0, 1fr);
	}
}

.pxh-preview {
	position: relative;
	border: 1px solid var(--vp-c-divider);
	border-radius: 8px;
	overflow: hidden;
	background: #eceef2;
}

.pxh-preview iframe {
	display: block;
	width: 100%;
	height: 60vh;
	min-height: 320px;
	border: 0;
}

.pxh-empty {
	position: absolute;
	inset: 0;
	display: grid;
	place-items: center;
	margin: 0;
	font-size: 13px;
	color: #454a57;
}

.pxh-side {
	max-height: 60vh;
	overflow-y: auto;
}

.pxh-reread {
	font-size: 12px;
	border: 1px solid var(--vp-c-divider);
	border-radius: 6px;
	padding: 3px 10px;
}

.pxh-lanes {
	margin: 8px 0 0;
	padding: 0;
	list-style: none;
	font-size: 13px;
}

.pxh-lanes li {
	display: flex;
	align-items: center;
	gap: 8px;
	flex-wrap: wrap;
	padding: 4px 0;
}

.pxh-lane {
	font-size: 11px;
	text-transform: uppercase;
	letter-spacing: 0.06em;
	border-radius: 20px;
	padding: 1px 9px;
	border: 1px solid var(--vp-c-divider);
	cursor: help;
}

.pxh-lane.is-exact {
	border-color: var(--vp-c-success-1);
	color: var(--vp-c-success-1);
}

.pxh-lane.is-reconciled {
	border-color: var(--vp-c-brand-1);
	color: var(--vp-c-brand-1);
}

.pxh-lane.is-drifted,
.pxh-lane.is-heuristic {
	border-color: var(--vp-c-warning-1);
	color: var(--vp-c-warning-1);
}

.pxh-lane-slide {
	color: var(--vp-c-text-2);
}

.pxh-lane-edits,
.pxh-lane-note {
	font-size: 12px;
	color: var(--vp-c-text-3);
}

.pxh-warnings {
	margin-top: 12px;
	font-size: 13px;
}

.pxh-warnings ul {
	margin: 6px 0 0;
	font-size: 12.5px;
	line-height: 1.6;
	color: var(--vp-c-text-2);
}

.pxh-downloads {
	display: flex;
	gap: 10px;
	flex-wrap: wrap;
	margin-top: 16px;
}

.pxh-downloads button {
	border: 1px solid var(--vp-c-divider);
	border-radius: 8px;
	padding: 7px 14px;
	font-size: 13px;
	background: var(--vp-c-bg-soft);
}

.pxh-downloads .pxh-primary {
	border-color: var(--vp-c-brand-1);
	color: var(--vp-c-brand-1);
}
</style>
