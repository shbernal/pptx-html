<script setup lang="ts">
/**
 * The top of the landing page.
 *
 * Static markup and no dependency on the library, which is why it is registered
 * synchronously while the marquee below it is not: this is what a visitor sees
 * before a megabyte of deck-writing code has finished arriving, and it is
 * pre-rendered into the HTML rather than appearing when Vue wakes up.
 *
 * `withBase` on every link, because the site is served from `/pptx-html/` on
 * GitHub Pages. Markdown links are rewritten by VitePress; hand-written `href`s
 * are not, and a bare `/playground` here is a 404 in production and nowhere else.
 */

import { withBase } from 'vitepress'
import { ref } from 'vue'

/**
 * `pnpm` rather than `npm`, in the one place a visitor is most likely to copy a
 * line without reading it. It is the package manager this repository declares
 * and the one every command in `CONTRIBUTING.md` uses, so the front door should
 * not suggest a different one.
 */
const INSTALL = 'pnpm add pptx-html'

const copied = ref(false)

async function copy(): Promise<void> {
	try {
		await navigator.clipboard.writeText(INSTALL)
		copied.value = true
		setTimeout(() => {
			copied.value = false
		}, 1600)
	} catch {
		// A refused clipboard is not worth an error state: the command is on screen
		// and selectable, which is the fallback.
	}
}
</script>

<template>
	<header class="pxh-hero">
		<div class="pxh-hero-inner">
			<p class="pxh-eyebrow">
				<span class="pxh-dot"></span>
				HTML ⇄ PPTX · lossless by construction
			</p>

			<h1>
				Edit a PowerPoint deck<br />
				<em>as a web page.</em>
			</h1>

			<p class="pxh-lead">
				And get the deck back: not an approximation of it. <code>pptx-html</code> reads a <code>.pptx</code> into a
				slide model, renders it as HTML you can edit, reads the edit back and writes the deck out again. Four legs,
				one loop, no screenshots anywhere.
			</p>

			<div class="pxh-actions">
				<a class="pxh-cta is-primary" :href="withBase('/playground')">Run the loop in your browser</a>
				<a class="pxh-cta" :href="withBase('/docs/')">Read the design record</a>
			</div>

			<button type="button" class="pxh-install" @click="copy">
				<code>{{ INSTALL }}</code>
				<span>{{ copied ? 'copied' : 'copy' }}</span>
			</button>
		</div>
	</header>
</template>

<style scoped>
.pxh-hero {
	position: relative;
	padding: 84px 24px 46px;
	overflow: hidden;
	background:
		radial-gradient(1100px 460px at 12% -12%, rgb(61 90 254 / 16%), transparent 60%),
		radial-gradient(820px 380px at 88% 4%, rgb(216 30 91 / 12%), transparent 62%);
}

/* The hairline the eye reads as the edge of the header band. */
.pxh-hero::after {
	content: '';
	position: absolute;
	inset: auto 0 0;
	height: 1px;
	background: linear-gradient(to right, transparent, var(--vp-c-divider) 20%, var(--vp-c-divider) 80%, transparent);
}

.pxh-hero-inner {
	max-width: 1152px;
	margin: 0 auto;
}

.pxh-eyebrow {
	display: inline-flex;
	align-items: center;
	gap: 8px;
	margin: 0 0 22px;
	padding: 5px 13px 5px 9px;
	border: 1px solid var(--vp-c-divider);
	border-radius: 20px;
	background: var(--vp-c-bg-soft);
	font-size: 12px;
	font-weight: 500;
	letter-spacing: 0.02em;
	color: var(--vp-c-text-2);
}

.pxh-dot {
	width: 7px;
	height: 7px;
	border-radius: 50%;
	background: var(--vp-c-brand-1);
}

.pxh-hero h1 {
	margin: 0;
	font-size: clamp(38px, 6.2vw, 68px);
	line-height: 1.06;
	letter-spacing: -0.028em;
	font-weight: 700;
	color: var(--vp-c-text-1);
}

.pxh-hero h1 em {
	font-style: normal;
	background: linear-gradient(100deg, var(--vp-c-brand-1), #8a7cff 46%, #00c2cb);
	background-clip: text;
	-webkit-background-clip: text;
	color: transparent;
}

.pxh-lead {
	max-width: 640px;
	margin: 24px 0 0;
	font-size: 16.5px;
	line-height: 1.66;
	color: var(--vp-c-text-2);
}

.pxh-lead code {
	font-size: 14.5px;
	padding: 2px 5px;
	border-radius: 5px;
	background: var(--vp-c-bg-soft);
}

.pxh-actions {
	display: flex;
	flex-wrap: wrap;
	gap: 12px;
	margin-top: 30px;
}

.pxh-cta {
	display: inline-block;
	padding: 11px 22px;
	border: 1px solid var(--vp-c-divider);
	border-radius: 10px;
	font-size: 14.5px;
	font-weight: 600;
	color: var(--vp-c-text-1);
	background: var(--vp-c-bg-soft);
	transition:
		border-color 0.2s,
		transform 0.2s;
}

.pxh-cta:hover {
	border-color: var(--vp-c-brand-1);
	transform: translateY(-1px);
}

.pxh-cta.is-primary {
	border-color: transparent;
	color: #fff;
	background: linear-gradient(100deg, var(--vp-c-brand-1), #6c4bf6);
}

.pxh-install {
	display: inline-flex;
	align-items: center;
	gap: 12px;
	margin-top: 22px;
	padding: 7px 8px 7px 14px;
	border: 1px dashed var(--vp-c-divider);
	border-radius: 9px;
	background: transparent;
	color: var(--vp-c-text-2);
}

.pxh-install code {
	font-size: 13px;
	color: var(--vp-c-text-1);
}

.pxh-install span {
	padding: 2px 9px;
	border-radius: 6px;
	background: var(--vp-c-bg-soft);
	font-size: 11px;
	letter-spacing: 0.06em;
	text-transform: uppercase;
	color: var(--vp-c-text-3);
}

.pxh-install:hover {
	border-color: var(--vp-c-brand-1);
}

@media (max-width: 640px) {
	.pxh-hero {
		padding: 56px 22px 36px;
	}
}
</style>
