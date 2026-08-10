<script setup lang="ts">
import type { Color, EditableParaProp, EditableRunProp } from 'pptx-html'
import {
	ALIGN_CHOICES,
	BULLET_CHOICES,
	BULLET_STATED,
	bulletChoiceOf,
	bulletValueOf,
	CONTROLS,
	DECORATION_CHOICES,
	PARA_CONTROLS,
	type ParagraphRow,
	POINTS_BOUNDS,
	type RunRow,
	type SlideRow,
	srgb,
} from './surface.ts'

defineProps<{ slides: SlideRow[] }>()

/**
 * One event for every property, carrying the key rather than encoding it in the
 * event name. `undefined` is *clear the key*, which is how this model spells
 * inherited — so the panel can express the difference between "not struck
 * through" and "says nothing about being struck through".
 *
 * `paraProp` is a second event rather than a flag on the first because the two
 * address different things — `node/paragraph/run` against `node/paragraph` — and
 * a single event would have to be told which, which is a discriminator waiting to
 * be passed wrongly.
 */
const emit = defineEmits<{
	text: [address: string, value: string]
	prop: [address: string, prop: EditableRunProp, value: unknown]
	paraProp: [address: string, prop: EditableParaProp, value: unknown]
	remove: [id: string]
}>()

/** The hex to put in the colour input. A scheme colour shows what it resolved to. */
function swatch(color: Color | undefined): string {
	if (color === undefined) return '#000000'
	return `#${color.kind === 'srgb' ? color.hex : color.effectiveHex}`
}

function isScheme(run: RunRow): boolean {
	return run.props.color?.kind === 'scheme'
}

function onSize(address: string, value: string) {
	const parsed = Number.parseFloat(value)
	emit('prop', address, 'sizePt', value.trim() === '' || Number.isNaN(parsed) ? undefined : parsed)
}

/** The empty option is *inherited*; every other value is one the run states outright. */
function onDecoration(address: string, prop: EditableRunProp, value: string) {
	emit('prop', address, prop, value === '' ? undefined : value)
}

/** The same reading of the empty option, one tier up. */
function onAlign(address: string, prop: EditableParaProp, value: string) {
	emit('paraProp', address, prop, value === '' ? undefined : value)
}

/**
 * The bullet select's positions are not its values: `''` clears the key, `none`
 * is the explicit `a:buNone`, and `bullet` is a glyph. The fourth position is
 * `disabled`, so a change event can never carry it.
 */
function onBullet(address: string, value: string) {
	emit('paraProp', address, 'bullet', bulletValueOf(value))
}

function onColor(address: string, value: string) {
	emit('prop', address, 'color', srgb(value))
}

/**
 * A margin in points, or the empty field that clears it.
 *
 * Empty is *inherited* here for the same reason the empty `<option>` is on the
 * drop-downs — and it is the position with something to say: a paragraph whose
 * margin is cleared follows its list style again, which is a different paragraph
 * from one stating `0`. Both are reachable from this input.
 */
function onPoints(address: string, prop: EditableParaProp, value: string) {
	const parsed = Number.parseFloat(value)
	emit('paraProp', address, prop, value.trim() === '' || Number.isNaN(parsed) ? undefined : parsed)
}

function pointsBounds(prop: EditableParaProp): { min: number; max: number } {
	return POINTS_BOUNDS[prop as 'marginLeftPt' | 'indentPt']
}

/** What to show in the field: the stated points, or nothing when the paragraph inherits. */
function pointsValue(props: ParagraphRow['props'], prop: EditableParaProp): number | '' {
	const value = props[prop as 'marginLeftPt' | 'indentPt']
	return typeof value === 'number' ? value : ''
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

				<div v-for="para in node.paragraphs" :key="para.address" class="pxh-para">
					<div class="pxh-para-head">
						<span class="pxh-para-label">paragraph</span>
						<label v-for="control in PARA_CONTROLS" :key="control.prop">
							{{ control.prop }}
							<select
								v-if="control.kind === 'align'"
								class="pxh-select"
								:value="para.props.align ?? ''"
								@change="onAlign(para.address, control.prop, ($event.target as HTMLSelectElement).value)"
							>
								<option v-for="choice in ALIGN_CHOICES" :key="choice.value" :value="choice.value">
									{{ choice.label }}
								</option>
							</select>
							<select
								v-else-if="control.kind === 'bullet'"
								class="pxh-select"
								:value="bulletChoiceOf(para.props.bullet)"
								@change="onBullet(para.address, ($event.target as HTMLSelectElement).value)"
							>
								<option
									v-for="choice in BULLET_CHOICES"
									:key="choice.value"
									:value="choice.value"
									:disabled="choice.value === BULLET_STATED"
								>
									{{ choice.label }}
								</option>
							</select>
							<input
								v-else
								class="pxh-size"
								type="number"
								step="1"
								:min="pointsBounds(control.prop).min"
								:max="pointsBounds(control.prop).max"
								:value="pointsValue(para.props, control.prop)"
								placeholder="inherited"
								@change="onPoints(para.address, control.prop, ($event.target as HTMLInputElement).value)"
							/>
						</label>
					</div>

					<p v-if="para.runs.length === 0" class="pxh-empty">
						A blank line. It states no text to edit, and its paragraph properties are editable like any other's.
					</p>

					<div v-for="run in para.runs" :key="run.address" class="pxh-run">
					<input
						class="pxh-text"
						type="text"
						:value="run.text"
						:aria-label="`Text of run ${run.address}`"
						@change="emit('text', run.address, ($event.target as HTMLInputElement).value)"
					/>
					<div class="pxh-props">
						<template v-for="control in CONTROLS" :key="control.prop">
							<label v-if="control.kind === 'toggle'">
								<input
									type="checkbox"
									:checked="run.props[control.prop] === true"
									@change="
										emit('prop', run.address, control.prop, ($event.target as HTMLInputElement).checked ? true : undefined)
									"
								/>
								{{ control.prop }}
							</label>
							<label v-else-if="control.kind === 'decoration'">
								{{ control.prop }}
								<select
									class="pxh-select"
									:value="run.props[control.prop] ?? ''"
									@change="onDecoration(run.address, control.prop, ($event.target as HTMLSelectElement).value)"
								>
									<option v-for="choice in DECORATION_CHOICES" :key="choice.value" :value="choice.value">
										{{ choice.label }}
									</option>
								</select>
							</label>
							<label v-else-if="control.kind === 'number'">
								{{ control.prop }}
								<input
									class="pxh-size"
									type="number"
									min="1"
									step="1"
									:value="run.props.sizePt ?? ''"
									placeholder="inherited"
									@change="onSize(run.address, ($event.target as HTMLInputElement).value)"
								/>
							</label>
							<label v-else>
								{{ control.prop }}
								<input
									type="color"
									:value="swatch(run.props.color)"
									@change="onColor(run.address, ($event.target as HTMLInputElement).value)"
								/>
							</label>
						</template>
						<button v-if="run.props.color" type="button" class="pxh-clear" @click="emit('prop', run.address, 'color', undefined)">
							clear
						</button>
					</div>
					<p v-if="isScheme(run)" class="pxh-scheme">
						This run states a <strong>scheme</strong> colour: a reference that re-resolves against the theme.
						Setting a value here replaces the reference with a fixed one. That is a legal edit; it is just not
						the same fact.
					</p>
					</div>
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

.pxh-para {
	margin-top: 8px;
	padding-left: 8px;
	border-left: 2px solid var(--vp-c-divider);
}

.pxh-para-head {
	display: flex;
	align-items: center;
	gap: 10px;
	flex-wrap: wrap;
	font-size: 12px;
	color: var(--vp-c-text-2);
}

.pxh-para-label {
	font-size: 11px;
	text-transform: uppercase;
	letter-spacing: 0.06em;
	color: var(--vp-c-text-3);
}

.pxh-para-head label {
	display: inline-flex;
	align-items: center;
	gap: 4px;
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

.pxh-select {
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
