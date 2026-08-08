import { defineConfig } from 'tsdown'

// Browser-only package (needs a real DOM at runtime). @shbernal/ts-pptx is a
// real dependency of this package, not bundled here; the consumer's bundler
// resolves it.
export default defineConfig({
	clean: true,
	dts: {
		sourcemap: true,
	},
	entry: {
		index: 'src/index.ts',
	},
	fixedExtension: false,
	format: 'esm',
	platform: 'browser',
	sourcemap: true,
	target: 'es2024',
	treeshake: true,
})
