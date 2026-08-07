/** Embedding contract with the Nextcloud hufak app.
 *
 * The app is framed cross-origin, so it inherits nothing from the host page.
 * The host hands over its resolved theme values, which are applied as CSS
 * custom properties; every style in this app resolves them with a standalone
 * fallback, so the same stylesheet serves both hosts.
 *
 * host -> app   #theme=<json> on the initial URL (applied before first paint)
 *               { type: 'hufak:theme', theme: Record<string, string> }
 * app  -> host  { type: 'hufak:ready' }
 *               { type: 'hufak:height', height: number }
 */

const HOST_ORIGIN_PATTERN = /^https:\/\/[a-z0-9.-]+\.hufak\.net$/;

/** The host origin is taken from the referrer so height reports are not
 * broadcast to any parent; theme messages are validated against it too. */
function resolveHostOrigin(): string | null {
	try {
		const origin = new URL(document.referrer).origin;
		return HOST_ORIGIN_PATTERN.test(origin) ? origin : null;
	} catch {
		return null;
	}
}

function isEmbedded(): boolean {
	return (
		window.self !== window.top
		&& new URLSearchParams(window.location.search).get('embed') === '1'
	);
}

function applyTheme(theme: unknown): void {
	if (!theme || typeof theme !== 'object') {
		return;
	}
	Object.entries(theme as Record<string, unknown>).forEach(([token, value]) => {
		// only ever set custom properties, never arbitrary CSS
		if (token.startsWith('--') && typeof value === 'string' && value.length < 200) {
			document.documentElement.style.setProperty(token, value);
		}
	});
}

function applyThemeFromHash(): void {
	const match = /(?:^|&)theme=([^&]+)/.exec(window.location.hash.slice(1));
	if (!match) {
		return;
	}
	try {
		applyTheme(JSON.parse(decodeURIComponent(match[1])));
	} catch {
		// a malformed hash just leaves the standalone palette in place
	}
}

function reportHeight(hostOrigin: string): void {
	const height = Math.max(
		document.documentElement.scrollHeight,
		document.body?.scrollHeight ?? 0,
	);
	window.parent.postMessage({ type: 'hufak:height', height }, hostOrigin);
}

/** Sets up theme adoption and height reporting when running inside the
 * Nextcloud app. Does nothing on the standalone deployment. */
function setupEmbedding(): void {
	if (!isEmbedded()) {
		return;
	}

	document.documentElement.dataset.embed = '1';
	applyThemeFromHash();

	const hostOrigin = resolveHostOrigin();
	if (!hostOrigin) {
		// framed by something other than the Nextcloud app: adopt nothing further
		return;
	}

	window.addEventListener('message', (event: MessageEvent) => {
		if (event.origin !== hostOrigin || event.source !== window.parent) {
			return;
		}
		const data = event.data as { type?: string; theme?: unknown } | null;
		if (data?.type === 'hufak:theme') {
			applyTheme(data.theme);
		}
	});

	const report = () => reportHeight(hostOrigin);
	new ResizeObserver(report).observe(document.documentElement);
	window.addEventListener('load', report);
	window.parent.postMessage({ type: 'hufak:ready' }, hostOrigin);
}

export { setupEmbedding };
