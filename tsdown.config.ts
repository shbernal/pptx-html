import { defineConfig } from 'tsdown'

// Browser-only package (needs a real DOM at runtime). @shbernal/ts-pptx and
// html2canvas are real dependencies of this package, not bundled here; the
// consumer's bundler resolves them.
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
