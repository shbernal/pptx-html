import { describe, expect, it } from 'vitest'

import {
	removeNoteRelationships,
	repairSlideNonVisualIds,
	repairSlideTextAutofit,
	repairSlideXml,
	slideNumber,
} from '../../src/repair/repair'

// The post-write OOXML repairs are pure string→string transforms, so they are
// unit-testable in Node without a DOM (unlike the rest of the engine).
// These lock in the behavior carried over from the original browser engine.

describe('repair (pure OOXML transforms)', () => {
	it('slideNumber parses the slide index from a part name', () => {
		expect(slideNumber('ppt/slides/slide12.xml')).toBe(12)
		expect(slideNumber('ppt/slides/slide1.xml')).toBe(1)
		expect(slideNumber('ppt/presentation.xml')).toBe(0)
	})

	it('repairSlideNonVisualIds de-duplicates colliding cNvPr ids', () => {
		const xml = '<p:cNvPr id="1" name="a"/><p:cNvPr id="1" name="b"/>'
		expect(repairSlideNonVisualIds(xml)).toBe('<p:cNvPr id="1" name="a"/><p:cNvPr id="2" name="b"/>')
	})

	it('repairSlideTextAutofit forces a single noAutofit on bodyPr', () => {
		expect(repairSlideTextAutofit('<a:bodyPr wrap="square"/>')).toBe(
			'<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>'
		)
		expect(repairSlideTextAutofit('<a:bodyPr><a:spAutoFit/></a:bodyPr>')).toBe('<a:bodyPr><a:noAutofit/></a:bodyPr>')
	})

	it('repairSlideXml retargets anchor="mid" and applies autofit repair', () => {
		expect(repairSlideXml('<a:bodyPr anchor="mid"/>')).toBe('<a:bodyPr anchor="ctr"><a:noAutofit/></a:bodyPr>')
	})

	it('removeNoteRelationships strips only note relationships', () => {
		const xml =
			'<Relationship Id="rId2" Target="../notesSlides/notesSlide1.xml"/>' +
			'<Relationship Id="rId3" Target="slide1.xml"/>'
		expect(removeNoteRelationships(xml)).toBe('<Relationship Id="rId3" Target="slide1.xml"/>')
	})
})
