/**
 * Read model → IR geometry.
 *
 * The shortest file here, and deliberately so: the read model already decodes
 * `a:custGeom` into a typed {@link GeometryCommand} list whose verbs are the ones
 * the write API's freeform DSL takes, so there is no second path vocabulary to
 * invent. `src/heuristic/svg-path.ts` has one of its own for the DOM lane, and part
 * 06 reconciles it onto this — not the other way round, because this one is the
 * format's.
 *
 * Coordinates stay in each path's own `0..w`/`0..h` unit space rather than being
 * scaled to EMU. That is how OOXML states them, it is what the emitter passes
 * through, and scaling here would mean every consumer had to know the shape's box
 * to read a path back out.
 */

import { type AnyShape, isAutoShape } from '@shbernal/ts-pptx/read'
import type { Geometry, GeometryPath } from '../ir/render'
import { type ImportScope, note } from './context'

/**
 * A shape with no geometry at all still has to be drawable, and a rectangle is
 * what PowerPoint itself falls back to. The note is what keeps that from being a
 * silent invention.
 */
const FALLBACK: Geometry = { kind: 'preset', preset: 'rect', adjustValues: {} }

export function geometryOf(shape: AnyShape, scope: ImportScope): Geometry {
	if (isAutoShape(shape)) {
		const custom = shape.customGeometry
		if (custom !== null) {
			const paths: GeometryPath[] = custom.paths.map((path) => ({
				w: path.w,
				h: path.h,
				fill: path.fill,
				stroke: path.stroke,
				commands: path.commands,
			}))
			if (paths.length > 0) return { kind: 'custom', paths }
		}
	}

	const preset = shape.presetGeometry
	if (preset !== null) return { kind: 'preset', preset, adjustValues: { ...shape.adjustValues } }

	note(
		scope,
		'shape.empty',
		'approximated',
		'unread',
		'this shape states neither a preset nor a custom geometry, so it is drawn as a plain rectangle'
	)
	return FALLBACK
}
