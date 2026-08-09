<script setup lang="ts">
import type { Stage } from './loop.ts'

defineProps<{ stages: Stage[] }>()

const WHAT: Record<string, string> = {
	importDeck: 'bytes in, paint model out',
	renderDeck: 'model to a document with a JSON island beside the picture',
	parseDeck: 'the document read back, one lane per slide',
	emitDeck: 'model plus the source package, back to .pptx',
}
</script>

<template>
	<ol class="pxh-strip">
		<li v-for="stage in stages" :key="stage.name" class="pxh-stage" :class="`is-${stage.status}`">
			<div class="pxh-stage-head">
				<code>{{ stage.name }}</code>
				<span v-if="stage.ms !== null" class="pxh-stage-ms">{{ Math.round(stage.ms) }} ms</span>
			</div>
			<p v-if="stage.error" class="pxh-stage-error">{{ stage.error }}</p>
			<p v-else-if="stage.note" class="pxh-stage-note">{{ stage.note }}</p>
			<p v-else class="pxh-stage-what">{{ WHAT[stage.name] }}</p>
		</li>
	</ol>
</template>

<style scoped>
.pxh-strip {
	display: grid;
	grid-template-columns: repeat(auto-fit, minmax(190px, 1fr));
	gap: 10px;
	margin: 0;
	padding: 0;
	list-style: none;
}

.pxh-stage {
	border: 1px solid var(--vp-c-divider);
	border-left-width: 4px;
	border-radius: 8px;
	padding: 10px 12px;
	background: var(--vp-c-bg-soft);
}

.pxh-stage-head {
	display: flex;
	align-items: baseline;
	justify-content: space-between;
	gap: 8px;
}

.pxh-stage code {
	font-size: 13px;
	background: none;
	padding: 0;
}

.pxh-stage-ms {
	font-size: 12px;
	color: var(--vp-c-text-3);
	font-variant-numeric: tabular-nums;
}

.pxh-stage p {
	margin: 4px 0 0;
	font-size: 12px;
	line-height: 1.45;
}

.pxh-stage-what {
	color: var(--vp-c-text-3);
}

.pxh-stage-note {
	color: var(--vp-c-text-2);
}

.pxh-stage-error {
	color: var(--vp-c-danger-1);
}

.is-idle {
	border-left-color: var(--vp-c-divider);
}

.is-running {
	border-left-color: var(--vp-c-brand-1);
}

.is-ok {
	border-left-color: var(--vp-c-success-1);
}

.is-failed {
	border-left-color: var(--vp-c-danger-1);
}

/* Not idle, and it matters that it does not look idle: this leg was never
   reached, because the one before it failed. */
.is-blocked {
	border-left-color: var(--vp-c-divider);
	opacity: 0.55;
}
</style>
