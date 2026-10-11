/**
 * The surface probe: one writer option, through the whole loop, judged on the XML.
 *
 * ```
 * newDeck + author(off) ──► baseline ┐
 *                                    ├─ signature: what the option wrote
 * newDeck + author(on)  ──► probe    ┘
 *                             │
 *                             ├─ html lane   (import → render → parse → emit)
 *                             └─ script lane (import → emit, both legs upstream)
 *                                    │
 *                     lost = signature facts the output no longer holds
 * ```
 *
 * The two lanes are what make the result a triage rather than a verdict. The
 * script lane is upstream code end to end, so a loss there is a `pptx-ts` issue;
 * a loss only the HTML lane takes is this repo's renderer, parser or emitter.
 *
 * The lanes are the oracle's own (`test/oracle/*-lane.ts`), and the oracle's
 * `diffDeckIr` report is kept alongside, so a probe can say both *the XML lost
 * this* and *the read model did / did not see it go*. When the XML loses a fact
 * and the report is clean, the probe has found the oracle's blind spot.
 */

import type TsPptx from 'pptx-ts'
import { newDeck } from '../corpus/decks'
import { htmlLoop } from '../oracle/html-lane'
import { roundTrip } from '../oracle/roundtrip'
import { scriptLoop } from '../oracle/script-lane'
import { lost, packageFacts, type Signature, signature } from './facts'

/** Why a loss is accepted without an upstream issue. */
type AcceptedKind =
	/** The option writes nothing distinguishable by design (a default, or authoring-time only). A writer that drops it is `unwritten`. */
	| 'inert'
	/** The output says the same thing in a different spelling. The reason must say how. */
	| 'equivalent'
	/** OOXML or the HTML tier cannot carry it, and no converter change would. */
	| 'unsupported'
	/** A loss in this repo's lane that is known and not yet fixed. Debt, counted as such. */
	| 'local'

export interface Probe {
	/** Unique, stable: the ledger keys on it. `addText.underline.style:dbl`. */
	id: string
	/** The inventory paths this probe exercises. */
	covers: string[]
	/** Build the deck; `on` is whether the option under test is applied. Nothing else may differ. */
	author: (pptx: TsPptx, on: boolean) => void | Promise<void>
	/** The `shbernal/pptx-ts` issue for a loss in the upstream legs. Required for one, forbidden without one. */
	issue?: string
	/**
	 * The `shbernal/pptx-ts` issue for a writer that drops this option, so the probe is
	 * inert until it is fixed. Kept apart from `issue` so the fix cannot pass unnoticed:
	 * once the option writes, the probe fails and points here, whatever the round trip
	 * then does with it.
	 */
	unwritten?: string
	/** A loss accepted without an issue, and why. */
	accepted?: { kind: AcceptedKind; reason: string }
}

type ProbeStatus =
	/** The option's signature survived the HTML lane. */
	| 'survives'
	/** The upstream-only lane already loses it: a `pptx-ts` issue. */
	| 'upstream'
	/** Only the HTML lane loses it: this repo's bug. */
	| 'local'
	/** The option wrote nothing, so the probe proves nothing. */
	| 'inert'
	/** A lane threw. */
	| 'crash'

export interface Observation {
	id: string
	covers: string[]
	status: ProbeStatus
	/** Whether the probe slide went through the authored lane or was carried verbatim. */
	lane: 'authored' | 'carried' | 'mixed'
	signature: Signature
	htmlLost: string[]
	scriptLost: string[]
	/** The oracle's view of the HTML lane, for spotting the blind spot. */
	undeclared: number
	declared: number
	error?: string
}

async function build(probe: Probe, on: boolean): Promise<Uint8Array> {
	const pptx = newDeck()
	await probe.author(pptx, on)
	return pptx.write({ outputType: 'uint8array' })
}

/** Run one probe through both lanes. Never throws: a crash is an observation. */
export async function observe(probe: Probe): Promise<Observation> {
	const base = { id: probe.id, covers: probe.covers }
	let sig: Signature = { added: [], removed: [] }
	try {
		const baseline = await build(probe, false)
		const probed = await build(probe, true)
		sig = signature(await packageFacts(baseline), await packageFacts(probed))
		const html = await roundTrip(probed, htmlLoop)
		const script = await roundTrip(probed, scriptLoop)
		const sources = new Set(html.input.ir.slides.map((slide) => slide.source))
		const lane = sources.size > 1 ? 'mixed' : sources.has('carried') ? 'carried' : 'authored'
		const htmlLost = lost(sig, await packageFacts(html.output.bytes))
		const scriptLost = lost(sig, await packageFacts(script.output.bytes))
		const empty = sig.added.length === 0 && sig.removed.length === 0
		const status: ProbeStatus = empty
			? 'inert'
			: htmlLost.length === 0
				? 'survives'
				: scriptLost.length > 0
					? 'upstream'
					: 'local'
		return {
			...base,
			status,
			lane,
			signature: sig,
			htmlLost,
			scriptLost,
			undeclared: html.report.undeclared.length,
			declared: html.report.declared.length,
		}
	} catch (error) {
		return {
			...base,
			status: 'crash',
			lane: 'authored',
			signature: sig,
			htmlLost: [],
			scriptLost: [],
			undeclared: 0,
			declared: 0,
			error: error instanceof Error ? `${error.name}: ${error.message}` : String(error),
		}
	}
}

const ISSUE_URL = /^https:\/\/github\.com\/shbernal\/pptx-ts\/issues\/\d+$/

/**
 * Whether an observation agrees with what the probe declares about itself, and if
 * not, what is wrong. `null` is agreement.
 *
 * Every status but `survives` must be accounted for, and `survives` must not be:
 * a pin left on a probe that now passes is a fix arriving unnoticed, which is the
 * failure `docs/round-trip.md` says a pinning test exists to prevent.
 */
export function disagreement(probe: Probe, seen: Observation): string | null {
	const { issue, unwritten, accepted } = probe
	if (issue && !ISSUE_URL.test(issue)) return `issue must be a shbernal/pptx-ts issue URL, got ${issue}`
	if (unwritten && !ISSUE_URL.test(unwritten)) return `unwritten must be a shbernal/pptx-ts issue URL, got ${unwritten}`
	if (unwritten && (issue || accepted))
		return '`unwritten` accounts for the probe on its own: drop `issue` and `accepted`'
	if (unwritten && seen.status !== 'inert')
		return `the option writes now (${seen.status}): remove \`unwritten\`, close ${unwritten} citing this probe, then classify the round trip`
	switch (seen.status) {
		case 'survives':
			if (issue) return `survives now: the pinned loss is gone. Remove \`issue\`, then close ${issue} citing this probe`
			if (accepted) return `survives now: remove \`accepted\` (${accepted.kind})`
			return null
		case 'inert':
			return unwritten || accepted?.kind === 'inert'
				? null
				: 'the option wrote nothing: pick a non-default value; if the writer drops it, file it and set `unwritten`; if it is authoring-time only, accept it as `inert` with a reason'
		case 'upstream':
			return issue || accepted?.kind === 'unsupported' || accepted?.kind === 'equivalent'
				? null
				: 'the upstream-only lane loses it: file a shbernal/pptx-ts issue and set `issue`'
		case 'local':
			return accepted?.kind === 'local' || accepted?.kind === 'equivalent' || accepted?.kind === 'unsupported'
				? null
				: 'only the HTML lane loses it: fix it here, or accept it with a reason'
		case 'crash':
			return issue || accepted ? null : `a lane threw (${seen.error}): fix it, or file it and set \`issue\``
	}
}
