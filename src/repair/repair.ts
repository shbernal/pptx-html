// @ts-nocheck
/**
 * Post-write OOXML repairs applied to the base64 PPTX the writer produces:
 * autofit normalisation, non-visual id de-duplication, generated-notes removal,
 * and slide-size typing. Ported verbatim from the original browser engine.
 *
 * The ZIP toolkit is ts-pptx's fflate-backed `@shbernal/ts-pptx/zip` export —
 * the same backend the writer uses — so dom2pptx carries no separate ZIP
 * dependency. Unlike JSZip's mutable archive object, that toolkit is functional:
 * `readZip` decompresses to a `path → bytes` map we mutate in place, then a
 * `ZipWriter` re-zips it. The individual XML transforms are pure string→string
 * functions and are unit-tested directly.
 */

import { readZip, ZipWriter } from '@shbernal/ts-pptx/zip'

const decoder = new TextDecoder()

export async function repairPptxBase64(base64) {
	if (!base64) return base64
	try {
		const entries = await readZip(base64ToBytes(base64))
		const slideNames = [...entries.keys()]
			.filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
			.sort((a, b) => slideNumber(a) - slideNumber(b))
		for (const name of slideNames) {
			const xml = entryText(entries, name)
			const repaired = repairSlideXml(xml)
			if (repaired !== xml) entries.set(name, repaired)
		}
		repairPresentationPackage(entries)
		const writer = new ZipWriter()
		for (const [name, data] of entries) writer.add(name, data)
		return await writer.generate('base64', { compression: true })
	} catch (e) {
		return base64
	}
}

export function slideNumber(name) {
	const m = String(name || '').match(/slide(\d+)\.xml$/i)
	return m ? parseInt(m[1], 10) : 0
}

export function repairSlideXml(xml) {
	return repairSlideTextAutofit(repairSlideNonVisualIds(xml)
		.replace(/\banchor="mid"/g, 'anchor="ctr"'))
}

export function repairSlideTextAutofit(xml) {
	return String(xml || '')
		.replace(/<a:bodyPr\b([^>]*?)\/>/g, '<a:bodyPr$1><a:noAutofit/></a:bodyPr>')
		.replace(/<a:bodyPr\b([^>]*)>([\s\S]*?)<\/a:bodyPr>/g, (match, attrs, body) => {
			const clean = String(body || '').replace(/<a:(?:spAutoFit|normAutofit|noAutofit)\b[^>]*\/>/g, '')
			return '<a:bodyPr' + attrs + '><a:noAutofit/>' + clean + '</a:bodyPr>'
		})
}

export function repairSlideNonVisualIds(xml) {
	const used = new Set()
	let next = 2
	return String(xml || '').replace(/<p:cNvPr\b([^>]*?)\bid="(\d+)"([^>]*)>/g, (match, before, id, after) => {
		let n = parseInt(id, 10)
		if (!Number.isFinite(n) || n < 1 || used.has(n)) {
			while (used.has(next)) next++
			n = next
		}
		used.add(n)
		if (n >= next) next = n + 1
		return '<p:cNvPr' + before + 'id="' + n + '"' + after + '>'
	})
}

function repairPresentationPackage(entries) {
	removeGeneratedNotes(entries)
	replaceZipText(entries, 'ppt/presentation.xml', (xml) => {
		let out = String(xml || '').replace(/<p:notesMasterIdLst>[\s\S]*?<\/p:notesMasterIdLst>/g, '')
		out = out.replace(/<p:sldSz\b([^>]*?)\/>/g, (match, attrs) => {
			if (/\btype=/.test(attrs)) return match
			return '<p:sldSz' + attrs + ' type="screen16x9"/>'
		})
		return out
	})
}

function removeGeneratedNotes(entries) {
	replaceZipText(entries, '[Content_Types].xml', (xml) => String(xml || '')
		.replace(/<Override\b[^>]*PartName="\/ppt\/notesSlides\/notesSlide\d+\.xml"[^>]*\/>/g, '')
		.replace(/<Override\b[^>]*PartName="\/ppt\/notesMasters\/notesMaster\d+\.xml"[^>]*\/>/g, ''))

	replaceZipText(entries, 'ppt/_rels/presentation.xml.rels', removeNoteRelationships)

	const relNames = [...entries.keys()]
		.filter((name) => /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/i.test(name))
	for (const name of relNames) replaceZipText(entries, name, removeNoteRelationships)

	removeZipPrefix(entries, 'ppt/notesSlides/')
	removeZipPrefix(entries, 'ppt/notesMasters/')
}

export function removeNoteRelationships(xml) {
	return String(xml || '').replace(/<Relationship\b(?=[^>]*\/notes(?:Slide|Master))[^>]*\/>/g, '')
}

function entryText(entries, name) {
	const data = entries.get(name)
	if (data == null) return undefined
	return typeof data === 'string' ? data : decoder.decode(data)
}

function replaceZipText(entries, name, transform) {
	if (!entries.has(name)) return
	const xml = entryText(entries, name)
	const updated = transform(xml)
	if (updated !== xml) entries.set(name, updated)
}

function removeZipPrefix(entries, prefix) {
	for (const name of [...entries.keys()]) {
		if (name.indexOf(prefix) === 0) entries.delete(name)
	}
}

function base64ToBytes(base64) {
	const binary = atob(base64)
	const bytes = new Uint8Array(binary.length)
	for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
	return bytes
}
