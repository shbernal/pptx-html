// Vite serves `?raw` imports as the file's text. Declared here so the browser
// fixtures (`import html from '../fixtures/x.html?raw'`) type-check.
declare module '*?raw' {
	const content: string
	export default content
}
