import { localName } from '../utils'
import { tokenize } from '../xml'

// Relationship parts (_rels/*.rels) are the source of truth for locating parts. A part's
// relationships live in `<dir>/_rels/<part>.rels`, and their Targets resolve against
// `<dir>` — so a worksheet is found by following workbook.xml's r:id into
// workbook.xml.rels, never by guessing a filename.

export interface Relationship {
	readonly id: string
	readonly type: string
	/** Target exactly as written; resolve internal ones with resolveTarget. */
	readonly target: string
	readonly targetMode: 'Internal' | 'External'
}

export function parseRels(xml: string): Map<string, Relationship> {
	const rels = new Map<string, Relationship>()
	for (const token of tokenize(xml)) {
		if (token.kind !== 'open' || localName(token.name) !== 'Relationship') continue
		const id = token.attrs.Id
		const target = token.attrs.Target
		if (id === undefined || target === undefined) continue
		rels.set(id, {
			id,
			type: token.attrs.Type ?? '',
			target,
			targetMode: token.attrs.TargetMode === 'External' ? 'External' : 'Internal',
		})
	}
	return rels
}

// Resolve a relationship Target (relative to `baseDir`, the directory of the owning part)
// to a package-absolute part path. A leading "/" makes the Target package-absolute; "."
// and ".." segments are honoured. External targets (URLs) should not be passed here.
export function resolveTarget(baseDir: string, target: string): string {
	const segments = target.startsWith('/') || baseDir === '' ? [] : baseDir.split('/')
	for (const part of target.split('/')) {
		if (part === '' || part === '.') continue
		if (part === '..') segments.pop()
		else segments.push(part)
	}
	return segments.join('/')
}
