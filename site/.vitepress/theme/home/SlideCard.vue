<script setup lang="ts">
/**
 * One rendered slide, sealed in a shadow root.
 *
 * The markup is `renderDeck`'s own output and goes in untouched — which is safe
 * for the reason the playground's frame is safe: every string the renderer takes
 * from a deck is escaped on the way out, so deck content is data here and never
 * markup. See `cardStyles` for why this is a shadow root rather than an iframe.
 */

import { onMounted, ref, watch } from 'vue'
import { cardStyles } from './showcase.ts'

const props = defineProps<{ markup: string; stylesheet: string }>()

const host = ref<HTMLElement | null>(null)

function install(): void {
	const element = host.value
	if (element === null) return
	// `attachShadow` throws if called twice on the same element, and the watcher
	// below re-runs whenever a row is re-rendered.
	const root = element.shadowRoot ?? element.attachShadow({ mode: 'open' })
	root.innerHTML = `<style>${cardStyles(props.stylesheet)}</style>${props.markup}`
}

onMounted(install)
watch(() => [props.markup, props.stylesheet], install)
</script>

<template>
	<div ref="host" class="pxh-card" aria-hidden="true"></div>
</template>

<style scoped>
/*
 * The frame is out here rather than in the shadow root: the elevation belongs to
 * the tile, and the slide inside it is the renderer's flat white page.
 */
.pxh-card {
	flex: 0 0 auto;
	width: var(--pxh-card-w);
	aspect-ratio: 16 / 9;
	margin-right: var(--pxh-card-gap);
	border-radius: 10px;
	overflow: hidden;
	background: #fff;
	box-shadow:
		0 1px 2px rgb(12 16 48 / 8%),
		0 12px 28px -12px rgb(12 16 48 / 35%);
	/* The row translates a strip several thousand pixels wide; without this the
	   compositor re-rasterizes every card on every frame. */
	transform: translateZ(0);
}
</style>
