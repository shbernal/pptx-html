/**
 * Assertions shared by all three test projects.
 *
 * `noUncheckedIndexedAccess` types every index read as possibly absent, which is
 * the truth: an out-of-range index is a defect, not an `undefined` to carry
 * forward. In a test the honest response is to fail at the read, naming what was
 * missing, rather than to optional-chain past it — a test that silently skips its
 * own assertion is worse than one that fails.
 */

/** The element at `index`, or a failure naming what was not there. */
export function at<T>(items: ArrayLike<T>, index: number, what = 'element'): T {
	const item = items[index]
	if (item === undefined) throw new Error(`no ${what} at index ${index}, in a list of ${items.length}`)
	return item
}

/** The first element, or a failure naming what was not there. */
export function first<T>(items: ArrayLike<T>, what = 'element'): T {
	return at(items, 0, what)
}

/** The only element, or a failure naming how many there actually were. */
export function only<T>(items: ArrayLike<T>, what = 'element'): T {
	if (items.length !== 1) throw new Error(`expected exactly one ${what}, found ${items.length}`)
	return at(items, 0, what)
}

/** `value`, or a failure when it is absent. */
export function present<T>(value: T | undefined | null, what = 'value'): T {
	if (value === undefined || value === null) throw new Error(`expected a ${what}, found none`)
	return value
}
