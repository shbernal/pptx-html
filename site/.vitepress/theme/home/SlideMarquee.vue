<script setup lang="ts">
/**
 * Two rows of slides, running in opposite directions.
 *
 * Everything in them is drawn by `renderDeck` in this tab when the component
 * mounts. There is no image fallback and there is deliberately nothing to fall
 * back to: if the loop cannot draw the decks, the row says so, because a picture
 * of a slide standing in for a rendered one is the exact claim this project
 * refuses to make.
 */

import { withBase } from 'vitepress'
import { computed, onMounted, ref, shallowRef } from 'vue'
import { download } from '../download.ts'
import { describe, kb } from '../playground/loop.ts'
import SlideCard from './SlideCard.vue'
import { showcase, type ShowcaseRow, type ShowcaseSlide } from './showcase.ts'

// Shallow because a row is written once and never mutated: deep reactivity would
// buy nothing and would walk sixteen slides' worth of markup to buy it.
const rows = shallowRef<ShowcaseRow[]>([])
const failure = ref<string | null>(null)

onMounted(async () => {
	try {
		rows.value = await showcase()
	} catch (error) {
		failure.value = describe(error)
	}
})

/**
 * The row's slides twice over.
 *
 * A track holding exactly two copies can be translated by half its width and
 * land on an identical frame, which is what makes the loop seamless. The gap is
 * a margin on each card rather than `gap` on the track for the same arithmetic:
 * with `gap` the track is `2n` cards and `2n − 1` gaps, and half of that is not
 * one set.
 */
function looped(row: ShowcaseRow): (ShowcaseSlide & { id: string })[] {
	return ['a', 'b'].flatMap((pass) => row.slides.map((slide) => ({ ...slide, id: `${slide.key}-${pass}` })))
}

const warnings = computed(() => rows.value.flatMap((row) => row.warnings))
</script>

<template>
	<div class="pxh-marquee">
		<p v-if="failure" class="pxh-marquee-failed">
			The decks could not be rendered in this browser, so there is nothing here: nothing on this site stands in
			for a slide it did not draw. <span>{{ failure }}</span>
		</p>

		<template v-else>
			<div
				v-for="(row, index) in rows"
				:key="row.name"
				class="pxh-row"
				:class="index % 2 === 0 ? 'is-leftward' : 'is-rightward'"
			>
				<div class="pxh-row-head">
					<span class="pxh-row-title">{{ row.title }}</span>
					<span class="pxh-row-blurb">{{ row.blurb }}</span>
					<button
						type="button"
						class="pxh-row-get"
						:aria-label="`Download ${row.title} as a .pptx file`"
						@click="download(row.source, row.file)"
					>
						<svg viewBox="0 0 16 16" width="12" height="12" aria-hidden="true">
							<path
								d="M8 1.5v8m0 0L4.75 6.25M8 9.5l3.25-3.25M2 11.5v1.75a1.25 1.25 0 0 0 1.25 1.25h9.5A1.25 1.25 0 0 0 14 13.25V11.5"
								fill="none"
								stroke="currentColor"
								stroke-width="1.5"
								stroke-linecap="round"
								stroke-linejoin="round"
							/>
						</svg>
						Download .pptx
						<span class="pxh-row-size">{{ kb(row.source.length) }}</span>
					</button>
				</div>
				<div class="pxh-viewport">
					<div class="pxh-track">
						<SlideCard
							v-for="card in looped(row)"
							:key="card.id"
							:markup="card.markup"
							:stylesheet="row.stylesheet"
						/>
					</div>
				</div>
			</div>

			<div v-if="rows.length === 0" class="pxh-row">
				<div class="pxh-viewport">
					<div class="pxh-track is-waiting">
						<div v-for="n in 8" :key="n" class="pxh-skeleton"></div>
					</div>
				</div>
				<p class="pxh-waiting">Writing two decks with <code>pptx-ts</code> and rendering them…</p>
			</div>
		</template>

		<p v-if="rows.length > 0" class="pxh-marquee-note">
			Sixteen slides, written by <code>pptx-ts</code> and drawn by <code>renderDeck</code> in this tab a
			moment ago. Not screenshots: this site has none, and
			<a :href="withBase('/docs/decisions')">that is a rule, not an oversight</a>.
			Each row's button hands you the <code>.pptx</code> its slides were drawn from: written in this tab, never
			fetched. The companies and figures are invented.
			<template v-if="warnings.length > 0">
				<br />
				<strong>{{ warnings.length }}</strong> render warning{{ warnings.length === 1 ? '' : 's' }}, shown rather
				than swallowed: <span class="pxh-warnings">{{ warnings.join(' · ') }}</span>
			</template>
		</p>
	</div>
</template>

<style scoped>
.pxh-marquee {
	--pxh-card-w: 340px;
	--pxh-card-gap: 20px;
	display: flex;
	flex-direction: column;
	gap: 26px;
}

.pxh-row-head {
	display: flex;
	align-items: baseline;
	gap: 12px;
	flex-wrap: wrap;
	max-width: 1152px;
	margin: 0 auto 8px;
	padding: 0 24px;
	width: 100%;
}

.pxh-row-title {
	font-size: 12.5px;
	font-weight: 600;
	letter-spacing: 0.04em;
	color: var(--vp-c-text-1);
}

.pxh-row-blurb {
	font-size: 12.5px;
	color: var(--vp-c-text-3);
}

/*
 * Pushed to the far end of the head, so the two rows' buttons line up with each
 * other and with the page's measure rather than trailing whatever length the
 * blurb happens to be.
 */
.pxh-row-get {
	margin-left: auto;
	display: inline-flex;
	align-items: center;
	gap: 6px;
	padding: 4px 11px;
	border: 1px solid var(--vp-c-divider);
	border-radius: 999px;
	background: var(--vp-c-bg-soft);
	color: var(--vp-c-text-2);
	font-size: 12px;
	font-weight: 500;
	line-height: 1.5;
	white-space: nowrap;
	cursor: pointer;
	transition:
		border-color 0.2s,
		color 0.2s;
}

.pxh-row-get:hover {
	border-color: var(--vp-c-brand-1);
	color: var(--vp-c-brand-1);
}

.pxh-row-size {
	color: var(--vp-c-text-3);
	font-variant-numeric: tabular-nums;
}

.pxh-row-get:hover .pxh-row-size {
	color: inherit;
}

/*
 * The rows run edge to edge — the marquee is the one block on the page with no
 * measure. Deliberately `100%` of a full-width parent rather than `100vw`, which
 * counts the scrollbar and would put a horizontal one under the whole site.
 */
.pxh-viewport {
	position: relative;
	overflow: hidden;
	/* The strip has no ends, so it must not appear to have any. */
	mask-image: linear-gradient(to right, transparent, #000 7%, #000 93%, transparent);
}

.pxh-track {
	display: flex;
	width: max-content;
	will-change: transform;
	animation: pxh-drift 68s linear infinite;
}

/* Same keyframe, run backwards: the second row drifts against the first. */
.is-rightward .pxh-track {
	animation-direction: reverse;
}

/* Long enough to read a slide, and it resumes where it stopped. */
.pxh-viewport:hover .pxh-track {
	animation-play-state: paused;
}

@keyframes pxh-drift {
	from {
		transform: translateX(0);
	}
	to {
		transform: translateX(-50%);
	}
}

.pxh-skeleton {
	flex: 0 0 auto;
	width: var(--pxh-card-w);
	aspect-ratio: 16 / 9;
	margin-right: var(--pxh-card-gap);
	border-radius: 10px;
	background: var(--vp-c-bg-soft);
	border: 1px solid var(--vp-c-divider);
}

.pxh-track.is-waiting {
	animation: none;
}

.pxh-waiting,
.pxh-marquee-note,
.pxh-marquee-failed {
	max-width: 1152px;
	margin: 0 auto;
	padding: 0 24px;
	font-size: 12.5px;
	line-height: 1.7;
	color: var(--vp-c-text-3);
}

.pxh-marquee-note code {
	font-size: 11.5px;
}

.pxh-marquee-note a {
	color: var(--vp-c-brand-1);
	text-decoration: underline;
	text-underline-offset: 2px;
}

.pxh-warnings {
	color: var(--vp-c-warning-1);
}

.pxh-marquee-failed {
	color: var(--vp-c-text-2);
}

.pxh-marquee-failed span {
	display: block;
	margin-top: 4px;
	color: var(--vp-c-danger-1);
}

@media (max-width: 720px) {
	.pxh-marquee {
		--pxh-card-w: 260px;
		--pxh-card-gap: 14px;
	}
}

/*
 * Motion is the decoration, not the content. With it turned off the rows become
 * an ordinary horizontal scroller — the same sixteen slides, moved by hand.
 */
@media (prefers-reduced-motion: reduce) {
	.pxh-track {
		animation: none;
	}

	.pxh-viewport {
		overflow-x: auto;
		mask-image: none;
	}
}
</style>
