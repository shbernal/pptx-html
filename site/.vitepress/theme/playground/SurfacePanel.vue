<script setup lang="ts">
import type { Color } from 'pptx-html'
import type { SlideRow } from './surface.ts'

defineProps<{ slides: SlideRow[] }>()

const emit = defineEmits<{
	text: [address: string, value: string]
	bold: [address: string, value: boolean]
	italic: [address: string, value: boolean]
	size: [address: string, value: number | null]
	color: [address: string, value: string | null]
	remove: [id: string]
}>()

/** The hex to put in the colour input. A scheme colour shows what it resolved to. */
function swatch(color: Color | null): string {
	if (color === null) return '#000000'
	return `#${color.kind === 'srgb' ? color.hex : color.effectiveHex}`
}

function isScheme(color: Color | null): boolean {
	return color !== null && color.kind === 'scheme'
}

function onSize(address: string, value: string) {
	const parsed = Number.parseFloat(value)
	emit('size', address, value.trim() === '' || Number.isNaN(parsed) ? null : parsed)
}
</script>

<template>
	<div class="pxh-panel">
		<section v-for="slide in slides" :key="slide.number" class="pxh-panel-slide">
			<h4>Slide {{ slide.number }}</h4>

			<div v-for="node in slide.nodes" :key="node.id" class="pxh-node">
				<div class="pxh-node-head">
					<code>{{ node.id }}</code>
					<span class="pxh-node-kind">{{ node.kind }}</span>
					<button type="button" class="pxh-remove" @click="emit('remove', node.id)">delete node</button>
				</div>

				<p v-if="node.runs.length === 0" class="pxh-empty">
					No text runs. Deleting the node is the only edit the surface defines for it.
				</p>

				<div v-for="run in node.runs" :key="run.address" class="pxh-run">
					<input
						class="pxh-text"
						type="text"
						:value="run.text"
						:aria-label="`Text of run ${run.address}`"
						@change="emit('text', run.address, ($event.target as HTMLInputElement).value)"
					/>
					<div class="pxh-props">
						<label>
							<input type="checkbox" :checked="run.bold" @change="emit('bold', run.address, ($event.target as HTMLInputElement).checked)" />
							bold
						</label>
						<label>
							<input type="checkbox" :checked="run.italic" @change="emit('italic', run.address, ($event.target as HTMLInputElement).checked)" />
							italic
						</label>
						<label>
							sizePt
							<input
								class="pxh-size"
								type="number"
								min="1"
								step="1"
								:value="run.sizePt ?? ''"
								placeholder="inherited"
								@change="onSize(run.address, ($event.target as HTMLInputElement).value)"
							/>
						</label>
						<label>
							color
							<input
								type="color"
								:value="swatch(run.color)"
								@change="emit('color', run.address, ($event.target as HTMLInputElement).value)"
							/>
						</label>
						<button v-if="run.color" type="button" class="pxh-clear" @click="emit('color', run.address, null)">
							clear
						</button>
					</div>
					<p v-if="isScheme(run.color)" class="pxh-scheme">
						This run states a <strong>scheme</strong> colour — a reference that re-resolves against the theme.
						Setting a value here replaces the reference with a fixed one. That is a legal edit; it is just not
						the same fact.
					</p>
				</div>
			</div>
		</section>
	</div>
</template>

<style scoped>
.pxh-panel {
	font-size: 13px;
}

.pxh-panel-slide h4 {
	margin: 16px 0 8px;
	font-size: 13px;
	text-transform: uppercase;
	letter-spacing: 0.06em;
	color: var(--vp-c-text-3);
}

.pxh-node {
	border: 1px solid var(--vp-c-divider);
	border-radius: 8px;
	padding: 8px 10px;
	margin-bottom: 8px;
	background: var(--vp-c-bg-soft);
}

.pxh-node-head {
	display: flex;
	align-items: center;
	gap: 8px;
	flex-wrap: wrap;
}

.pxh-node-head code {
	font-size: 12px;
	background: none;
	padding: 0;
}

.pxh-node-kind {
	font-size: 11px;
	color: var(--vp-c-text-3);
	border: 1px solid var(--vp-c-divider);
	border-radius: 20px;
	padding: 0 7px;
}

.pxh-remove {
	margin-left: auto;
	font-size: 11px;
	color: var(--vp-c-danger-1);
	border: 1px solid var(--vp-c-divider);
	border-radius: 6px;
	padding: 1px 8px;
}

.pxh-empty {
	margin: 6px 0 0;
	font-size: 12px;
	color: var(--vp-c-text-3);
}

.pxh-run + .pxh-run {
	margin-top: 10px;
	padding-top: 10px;
	border-top: 1px dashed var(--vp-c-divider);
}

.pxh-run {
	margin-top: 8px;
}

.pxh-text {
	width: 100%;
	border: 1px solid var(--vp-c-divider);
	border-radius: 6px;
	padding: 4px 7px;
	background: var(--vp-c-bg);
	color: var(--vp-c-text-1);
}

.pxh-props {
	display: flex;
	align-items: center;
	gap: 12px;
	flex-wrap: wrap;
	margin-top: 6px;
	font-size: 12px;
	color: var(--vp-c-text-2);
}

.pxh-props label {
	display: inline-flex;
	align-items: center;
	gap: 4px;
}

.pxh-size {
	width: 74px;
	border: 1px solid var(--vp-c-divider);
	border-radius: 6px;
	padding: 1px 5px;
	background: var(--vp-c-bg);
	color: var(--vp-c-text-1);
}

.pxh-clear {
	font-size: 11px;
	border: 1px solid var(--vp-c-divider);
	border-radius: 6px;
	padding: 1px 7px;
	color: var(--vp-c-text-3);
}

.pxh-scheme {
	margin: 6px 0 0;
	font-size: 11.5px;
	line-height: 1.5;
	color: var(--vp-c-text-3);
}
</style>
