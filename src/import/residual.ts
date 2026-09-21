/**
 * The residual channel: a slide's own XML, kept verbatim.
 *
 * This is what lets "best effort in what we model" coexist with a hard
 * round-trip guarantee. A slide holding a construct the write API cannot express
 * is not approximated and not dropped — its part crosses intact, and the nodes
 * beside it exist only so the page has something to show.
 *
 * **The trigger is upstream's verdict, not a second opinion.** `readModelToIr`
 * already decides which slides it cannot transcribe and marks them
 * `source: 'carried'`; re-deriving that rule here would give the renderer and the
 * emitter two different ideas of which slides are carried, and they would
 * disagree exactly on the slides where it matters most.
 *
 * Media is referenced by `r:id`, because that is how the XML refers to it. A
 * carried slide's `<a:blip r:embed="rId3"/>` means nothing without the map from
 * `rId3` to bytes, and `ExtractedSlide.media` — the shape this has to take on the
 * way back — is keyed the same way.
 */

import type { Slide } from '@shbernal/ts-pptx/read'
import type { Residual, ResidualAsset } from '../ir/render'
import type { AssetIndex } from './assets'

/** OPC relationship type for an embedded image. */
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

export function residualOf(slide: Slide, assets: AssetIndex): Residual {
	// The part's *original* bytes, not a reserialization: a carried slide is only
	// worth carrying if it crosses unchanged, and `Part.originalBytes` is the
	// untouched zip entry. ts-pptx 4.0 renamed it from `Part.bytes` precisely so
	// that a call site says which of the two it means; `serialize()` is the other
	// one, and it is the wrong one here.
	const xml = new TextDecoder().decode(slide.part.originalBytes)

	const referenced: ResidualAsset[] = []
	for (const relationship of slide.relationships.byType(IMAGE_REL)) {
		if (relationship.targetMode === 'External') continue
		const asset = assets.refFor(slide.relationships.resolveTarget(relationship.id))
		if (asset !== null) referenced.push({ relId: relationship.id, asset })
	}

	return { kind: 'slide', xml, assets: referenced }
}
