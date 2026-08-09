---
layout: page
pageClass: pxh-home
editLink: false
aside: false
title: 'Edit a PowerPoint deck as a web page'
description: 'pptx-html reads a .pptx into a slide model, renders it as editable HTML, reads the edit back and writes the deck out again — losslessly.'
---

<script setup>
// `withBase` because the site is served from `/pptx-html/` on GitHub Pages.
// VitePress rewrites markdown links for that automatically, which is why every
// link inside prose below is written as markdown — it is also what
// `ignoreDeadLinks: false` checks. The card grids are hand-written HTML, so they
// have to say it themselves.
import { withBase } from 'vitepress'
</script>

<HomeHero />

<ClientOnly>
	<SlideMarquee />
</ClientOnly>

<div class="pxh-page">

<section>

<p class="pxh-kicker">What it is</p>

## Four functions, arranged in a circle

`pptx-html` moves slides between HTML and PPTX by driving
[`@shbernal/ts-pptx`](https://www.npmjs.com/package/@shbernal/ts-pptx). It reads a
`.pptx` into a slide model, renders that model as HTML, reads the edited HTML back,
and writes a `.pptx` out again — so a deck can be edited by anything that can edit
a web page, and still be a deck afterwards.

<pre class="pxh-loop">.pptx  ──import──►  IR  ──render──►  HTML   (what a human sees / edits)
  ▲                  ▲                 │
  └────emit──────────┴─────parse───────┘   (what a machine reads back)</pre>

<div class="pxh-invariant">

**Invariant R.** For any deck this pipeline can write, `import → render → parse →
emit` produces a deck **equal under the normalized read model** to the input.
Slides whose features the IR does not model are carried across **byte-identical**
rather than approximated.

</div>

Equality is normalized rather than byte-for-byte: zip entry order, timestamps,
relationship ids and element ids all vary legally, and both sides are canonicalized
before diffing. [The property, and the harness that gates it →](/docs/round-trip)

</section>

<section>

<p class="pxh-kicker">The three states</p>

## Every construct is in one of three states, and never a fourth

<div class="pxh-grid">
	<div class="pxh-tile" style="--pxh-tile-tint: #3d5afe">
		<h3>Modeled</h3>
		<p>The intermediate representation represents it, and it survives the loop exactly.</p>
	</div>
	<div class="pxh-tile" style="--pxh-tile-tint: #00c2cb">
		<h3>Carried</h3>
		<p>The IR does not model it, so its XML moves across untouched. No approximation, and no loss.</p>
	</div>
	<div class="pxh-tile" style="--pxh-tile-tint: #f0a020">
		<h3>Warned</h3>
		<p>It can be neither modeled nor carried, and the conversion says so. A visible failure, never a silent one.</p>
	</div>
	<div class="pxh-tile is-ruled-out" style="--pxh-tile-tint: #c4405f">
		<h3>Approximated</h3>
		<p>The state that does not exist here: content that comes out looking about right and has no way back. It is why the raster fallback was removed rather than kept as an escape hatch — a slide flattened into a picture can never re-enter the loop — and why this site shows no pictures of slides.</p>
	</div>
</div>

Those three states are a claim, and the [fidelity ledger](/docs/fidelity) is its
evidence: what the loop currently models, carries and warns on, per deck, measured
from the same corpus the round-trip oracle gates on.

</section>

<section>

<p class="pxh-kicker">Scope</p>

## What the guarantee covers, and what it does not

A generated corpus of 17 decks runs the full loop on every CI build, and the
per-construct fidelity ledger is snapshotted so it cannot move silently. Claims
here are held to what that gate actually covers; where something is not covered,
it says so.

<ul class="pxh-facts">
	<li>
		<b>Input domain</b>
		<span>Decks written by <code>@shbernal/ts-pptx</code>, which is what the generated corpus is made of — and what the two decks moving above are written by. Decks authored in PowerPoint are a deliberate second tier and are <strong>not yet gated</strong>.</span>
	</li>
	<li>
		<b>Environment</b>
		<span>The loop is host-agnostic and runs in Node and the browser alike. The best-effort heuristic lane, for HTML this library did not render, is browser-only: it needs a real DOM.</span>
	</li>
	<li>
		<b>Distribution</b>
		<span>Published to npm as <code>pptx-html</code>. ESM only, Node <code>&gt;=24</code>. Pre-1.0: the loop's four legs are stable, the heuristic lane's model is not.</span>
	</li>
</ul>

</section>

<section>

<p class="pxh-kicker">Where to go next</p>

## Run it, or read why it is built this way

<div class="pxh-doors">
	<a class="pxh-door" :href="withBase('/playground')">
		<strong>Playground</strong>
		<span>All four legs, running on bytes that never leave your tab. Pick a corpus deck or drop one of your own, edit a run, then download both files and compare them.</span>
	</a>
	<a class="pxh-door" :href="withBase('/docs/')">
		<strong>The design record</strong>
		<span>Invariant R and the oracle, the two lanes and two models of the architecture, and the decisions that must not be undone.</span>
	</a>
	<a class="pxh-door" href="https://github.com/shbernal/pptx-html#readme">
		<strong>README and source</strong>
		<span>Install, the four-leg example, and the repository itself. MIT licensed.</span>
	</a>
</div>

</section>

</div>
