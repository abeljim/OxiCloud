import { it, expect, vi, beforeEach } from 'vitest';
import { render, screen, fireEvent, waitFor } from '@testing-library/svelte';
vi.mock('$lib/api/client', () => ({ apiFetch: vi.fn(), withBase: (p: string) => p }));
vi.mock('$lib/api/endpoints/files', () => ({
	fileDownloadUrl: () => '/dl',
	fileInlineUrl: () => '/in'
}));
vi.mock('$lib/api/endpoints/wopi', () => ({
	canEditWithWopi: vi.fn(),
	// WopiEditor's $effect calls this on mount when an editor opens; provide it
	// so the async handshake resolves instead of throwing an unhandled error
	// (vitest 4 fails the whole run on unhandled errors).
	getEditorUrlWithFallback: vi.fn(async () => ({
		editor_url: 'about:blank',
		access_token: 't',
		access_token_ttl: 0
	}))
}));
// Collab is off for most of these tests — a text file then takes the plain
// preview path the older assertions describe. The one collab test flips it.
const { cfg } = vi.hoisted(() => ({
	cfg: { loaded: true, features: { markdown_collab: false } }
}));
vi.mock('$lib/stores/serverConfig.svelte', () => ({ serverConfig: cfg }));
// Collab also needs a signed-in caller: it runs over the message bus, whose
// WS ticket route is deliberately off the anonymous share allowlist. These
// tests describe the authenticated case, so say so explicitly — see the
// anonymous test at the bottom for the other side.
// `load` is part of the contract the message bus relies on — it probes
// the session before asking for a WS ticket — so the stub needs it even
// though these tests never connect.
const { sess } = vi.hoisted(() => ({
	sess: { isAuthenticated: true, load: async () => null }
}));
vi.mock('$lib/stores/session.svelte', () => ({ session: sess }));
// The real editor mounts CodeMirror and a message-bus session; the stub just
// renders the toolbar the viewer passes in.
vi.mock(
	'$lib/components/CollabEditor.svelte',
	async () => await import('./__mocks__/CollabEditor.svelte')
);
import { apiFetch } from '$lib/api/client';
import { canEditWithWopi } from '$lib/api/endpoints/wopi';
import FileViewer from './FileViewer.svelte';
const af = apiFetch as unknown as ReturnType<typeof vi.fn>;
const cw = canEditWithWopi as unknown as ReturnType<typeof vi.fn>;
function file(over: Record<string, unknown> = {}) {
	return {
		id: 'i',
		name: 'pic.png',
		mime_type: 'image/png',
		category: 'Image',
		folder_id: '',
		created_by: null,
		updated_by: null,
		path: '',
		size: 1,
		modified_at: 0,
		created_at: 0,
		sort_date: 0,
		icon_class: '',
		icon_special_class: '',
		size_formatted: '1 B',
		etag: '',
		content_hash: '',
		...over
	} as never;
}
beforeEach(() => {
	vi.clearAllMocks();
	cw.mockResolvedValue(false);
});
it('renders an image with working zoom controls and closes', async () => {
	render(FileViewer, { props: { open: true, file: file() } });
	expect(await screen.findByTestId('file-viewer-dialog')).toBeTruthy();
	await fireEvent.click(screen.getByTestId('file-viewer-zoom-in-btn'));
	await fireEvent.click(screen.getByTestId('file-viewer-zoom-out-btn'));
	await fireEvent.click(screen.getByTestId('file-viewer-zoom-reset-btn'));
	await fireEvent.click(screen.getByTestId('file-viewer-close-btn'));
});
it('fetches text content for a text file', async () => {
	af.mockResolvedValue({ ok: true, text: async () => 'hello world' });
	render(FileViewer, {
		props: { open: true, file: file({ name: 'n.txt', mime_type: 'text/plain', category: 'Text' }) }
	});
	expect(await screen.findByTestId('file-viewer-dialog')).toBeTruthy();
	await waitFor(() => expect(af).toHaveBeenCalled());
});
it('renders nothing when closed', () => {
	render(FileViewer, { props: { open: false, file: file() } });
	expect(screen.queryByTestId('file-viewer-dialog')).toBeNull();
});
it('shows an Edit button for a WOPI-editable document', async () => {
	cw.mockResolvedValue(true);
	render(FileViewer, {
		props: {
			open: true,
			file: file({
				name: 'report.docx',
				mime_type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
				category: 'Document'
			})
		}
	});
	await screen.findByTestId('file-viewer-dialog');
	await waitFor(() => expect(screen.getByTestId('file-viewer-edit-btn')).toBeTruthy());
});
it('exposes download and open-in-new-tab links', async () => {
	render(FileViewer, { props: { open: true, file: file() } });
	await screen.findByTestId('file-viewer-dialog');
	expect(screen.getByTestId('file-viewer-download-link').getAttribute('href')).toBe('/dl');
	expect(screen.getByTestId('file-viewer-open-new-tab-link').getAttribute('href')).toBe('/in');
});
it('handles a failed text fetch without crashing', async () => {
	af.mockResolvedValue({ ok: false, status: 500, text: async () => '' });
	render(FileViewer, {
		props: { open: true, file: file({ name: 'n.txt', mime_type: 'text/plain', category: 'Text' }) }
	});
	await screen.findByTestId('file-viewer-dialog');
	await waitFor(() => expect(af).toHaveBeenCalled());
});

// A text-shaped file is edited in place, so the viewer hands its chrome to the
// editor: title + close stay on top, the actions move down next to the sync
// state — the shape the office editor already has.
it('moves the actions into the editor row for a collab-editable file', async () => {
	cfg.features.markdown_collab = true;
	render(FileViewer, {
		props: {
			open: true,
			file: file({ name: 'notes.md', mime_type: 'text/markdown', category: 'Document' })
		}
	});
	await screen.findByTestId('collab-editor-stub');
	expect(screen.getByTestId('file-viewer-collab-window-link')).toBeTruthy();
	expect(screen.getByTestId('file-viewer-collab-download-link')).toBeTruthy();
	// The top bar keeps only the title and the close button.
	expect(screen.queryByTestId('file-viewer-download-link')).toBeNull();
	expect(screen.queryByTestId('file-viewer-open-new-tab-link')).toBeNull();
	expect(screen.getByTestId('file-viewer-close-btn')).toBeTruthy();
	expect(screen.getByTestId('file-viewer-close-btn')).toBeTruthy();
	cfg.features.markdown_collab = false;
});

// A public-share visitor holds an anonymous session, and anonymous
// sessions never get a WebSocket — `POST /api/rt/ticket` is off their
// allowlist by design. Mounting the collaborative editor for them
// produced a 403 per attempt (`authz.denied`,
// `reason="anonymous_route_not_allowlisted"`) and an editor that could
// never sync. The plain preview is the correct surface: static, no
// realtime updates, which is all a share visitor can have.
it('shows the plain preview instead of the editor for an anonymous visitor', async () => {
	cfg.features.markdown_collab = true;
	sess.isAuthenticated = false;
	af.mockResolvedValue({ ok: true, text: async () => 'hello from the share' });

	render(FileViewer, {
		props: {
			open: true,
			file: file({ name: 'notes.md', mime_type: 'text/markdown', category: 'Document' })
		}
	});

	await screen.findByText('hello from the share');
	expect(screen.queryByTestId('collab-editor-stub')).toBeNull();

	sess.isAuthenticated = true;
	cfg.features.markdown_collab = false;
});

// `readOnly` is a separate gate from authentication: a host that opens
// the viewer read-only is asking for a preview, and that must hold even
// for a signed-in caller who could otherwise edit.
it('shows the plain preview instead of the editor when opened read-only', async () => {
	cfg.features.markdown_collab = true;
	af.mockResolvedValue({ ok: true, text: async () => 'read only body' });

	render(FileViewer, {
		props: {
			open: true,
			readOnly: true,
			file: file({ name: 'notes.md', mime_type: 'text/markdown', category: 'Document' })
		}
	});

	await screen.findByText('read only body');
	expect(screen.queryByTestId('collab-editor-stub')).toBeNull();

	cfg.features.markdown_collab = false;
});

// Fullscreen is the default, and the toggle sticks for the next file opened
// on this device.
it('remembers the editor fullscreen choice', async () => {
	cfg.features.markdown_collab = true;
	const md = file({ name: 'notes.md', mime_type: 'text/markdown', category: 'Document' });
	const first = render(FileViewer, { props: { open: true, file: md } });
	const toggle = await screen.findByTestId('file-viewer-collab-fullscreen-btn');
	expect(toggle.getAttribute('aria-pressed')).toBe('true');
	await fireEvent.click(toggle);
	expect(toggle.getAttribute('aria-pressed')).toBe('false');
	first.unmount();

	render(FileViewer, { props: { open: true, file: md } });
	const again = await screen.findByTestId('file-viewer-collab-fullscreen-btn');
	expect(again.getAttribute('aria-pressed')).toBe('false');
	localStorage.removeItem('oxi-collab-fullscreen');
	cfg.features.markdown_collab = false;
});

// Previous / next step through the images and videos of the list the viewer
// was opened from. Looking through a folder of photos used to mean closing
// the viewer and opening each one by hand.
const folder = () => [
	file({ id: 'a', name: 'a.jpg', mime_type: 'image/jpeg' }),
	file({ id: 'doc', name: 'report.docx', mime_type: 'application/octet-stream' }),
	file({ id: 'b', name: 'b.mp4', mime_type: 'video/mp4', category: 'Video' }),
	file({ id: 'c', name: 'c.png' })
];

it('steps to the neighbouring media with the previous / next buttons', async () => {
	const files = folder();
	const onnavigate = vi.fn();
	render(FileViewer, { props: { open: true, file: files[2], files, onnavigate } });
	await screen.findByTestId('file-viewer-dialog');

	// The .docx between a.jpg and b.mp4 is skipped.
	await fireEvent.click(screen.getByTestId('file-viewer-prev-btn'));
	expect(onnavigate).toHaveBeenLastCalledWith(files[0]);
	await fireEvent.click(screen.getByTestId('file-viewer-next-btn'));
	expect(onnavigate).toHaveBeenLastCalledWith(files[3]);
});

it('offers no step past either end', async () => {
	const files = folder();
	const first = render(FileViewer, {
		props: { open: true, file: files[0], files, onnavigate: vi.fn() }
	});
	await screen.findByTestId('file-viewer-dialog');
	expect(screen.queryByTestId('file-viewer-prev-btn')).toBeNull();
	expect(screen.getByTestId('file-viewer-next-btn')).toBeTruthy();
	first.unmount();

	render(FileViewer, { props: { open: true, file: files[3], files, onnavigate: vi.fn() } });
	await screen.findByTestId('file-viewer-dialog');
	expect(screen.getByTestId('file-viewer-prev-btn')).toBeTruthy();
	expect(screen.queryByTestId('file-viewer-next-btn')).toBeNull();
});

it('steps with the arrow keys, unless a focused player wants them', async () => {
	const files = folder();
	const onnavigate = vi.fn();
	render(FileViewer, { props: { open: true, file: files[2], files, onnavigate } });
	await screen.findByTestId('file-viewer-dialog');

	await fireEvent.keyDown(window, { key: 'ArrowRight' });
	expect(onnavigate).toHaveBeenLastCalledWith(files[3]);
	await fireEvent.keyDown(window, { key: 'ArrowLeft' });
	expect(onnavigate).toHaveBeenLastCalledWith(files[0]);
	// Alt+← is the browser's Back.
	await fireEvent.keyDown(window, { key: 'ArrowLeft', altKey: true });
	expect(onnavigate).toHaveBeenCalledTimes(2);

	// A focused <video> seeks with the arrows.
	const video = document.querySelector('video');
	expect(video).toBeTruthy();
	await fireEvent.keyDown(video!, { key: 'ArrowRight' });
	expect(onnavigate).toHaveBeenCalledTimes(2);
});

it('shows no stepping for a non-media file, or when the host does not handle it', async () => {
	const files = folder();
	cw.mockResolvedValue(false);
	const doc = render(FileViewer, {
		props: { open: true, file: files[1], files, onnavigate: vi.fn() }
	});
	await screen.findByTestId('file-viewer-dialog');
	expect(screen.queryByTestId('file-viewer-prev-btn')).toBeNull();
	expect(screen.queryByTestId('file-viewer-next-btn')).toBeNull();
	doc.unmount();

	render(FileViewer, { props: { open: true, file: files[2], files } });
	await screen.findByTestId('file-viewer-dialog');
	expect(screen.queryByTestId('file-viewer-prev-btn')).toBeNull();
	expect(screen.queryByTestId('file-viewer-next-btn')).toBeNull();
});

// Fullscreen takes the viewer's panel, not the page, so the previous / next
// controls come along and stepping keeps working. jsdom has no Fullscreen
// API, so these stub the part of it the viewer uses.
function stubFullscreen() {
	let current: Element | null = null;
	const enter = (el: Element) => {
		current = el;
		document.dispatchEvent(new Event('fullscreenchange'));
		return Promise.resolve();
	};
	const request = vi.fn(function (this: Element) {
		return enter(this);
	});
	const exit = vi.fn(() => {
		current = null;
		document.dispatchEvent(new Event('fullscreenchange'));
		return Promise.resolve();
	});
	Object.defineProperty(document, 'fullscreenEnabled', { configurable: true, value: true });
	Object.defineProperty(document, 'fullscreenElement', { configurable: true, get: () => current });
	Object.defineProperty(document, 'exitFullscreen', { configurable: true, value: exit });
	Object.defineProperty(HTMLElement.prototype, 'requestFullscreen', {
		configurable: true,
		value: request
	});
	return {
		request,
		exit,
		restore() {
			for (const key of ['fullscreenEnabled', 'fullscreenElement', 'exitFullscreen'] as const) {
				delete (document as unknown as Record<string, unknown>)[key];
			}
			delete (HTMLElement.prototype as unknown as Record<string, unknown>).requestFullscreen;
		}
	};
}

it('goes fullscreen with the previous / next controls, and keeps stepping there', async () => {
	const fs = stubFullscreen();
	const files = folder();
	const onnavigate = vi.fn();
	render(FileViewer, { props: { open: true, file: files[2], files, onnavigate } });
	await screen.findByTestId('file-viewer-dialog');

	const toggle = screen.getByTestId('file-viewer-fullscreen-btn');
	expect(toggle.getAttribute('aria-pressed')).toBe('false');
	await fireEvent.click(toggle);
	const target = fs.request.mock.contexts[0] as HTMLElement;
	expect(target.contains(screen.getByTestId('file-viewer-next-btn'))).toBe(true);
	expect(target.contains(toggle)).toBe(true);
	await waitFor(() => expect(toggle.getAttribute('aria-pressed')).toBe('true'));

	await fireEvent.keyDown(window, { key: 'ArrowRight' });
	expect(onnavigate).toHaveBeenLastCalledWith(files[3]);
	// Escape is the browser's way out of fullscreen; it must not also close.
	await fireEvent.keyDown(window, { key: 'Escape' });
	expect(screen.getByTestId('file-viewer-dialog')).toBeTruthy();

	await fireEvent.click(toggle);
	expect(fs.exit).toHaveBeenCalled();
	await waitFor(() => expect(toggle.getAttribute('aria-pressed')).toBe('false'));
	fs.restore();
});

it('offers no fullscreen where the browser cannot do it', async () => {
	render(FileViewer, { props: { open: true, file: file() } });
	await screen.findByTestId('file-viewer-dialog');
	expect(screen.queryByTestId('file-viewer-fullscreen-btn')).toBeNull();
});

it('shows previous / next while the pointer moves, and fades them once it rests', async () => {
	vi.useFakeTimers();
	const files = folder();
	render(FileViewer, { props: { open: true, file: files[2], files, onnavigate: vi.fn() } });
	const next = await screen.findByTestId('file-viewer-next-btn');
	const panel = next.parentElement!;
	expect(panel.classList.contains('fv__panel--nav')).toBe(false);

	await fireEvent.pointerMove(panel);
	expect(panel.classList.contains('fv__panel--nav')).toBe(true);
	await vi.advanceTimersByTimeAsync(2100);
	expect(panel.classList.contains('fv__panel--nav')).toBe(false);

	await fireEvent.pointerMove(panel);
	await fireEvent.pointerLeave(panel);
	expect(panel.classList.contains('fv__panel--nav')).toBe(false);
	vi.useRealTimers();
});
