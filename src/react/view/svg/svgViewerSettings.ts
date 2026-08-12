import { vscodeApi } from '../../util/vscode';

const LINE_WRAP_KEY = 'office-svg-line-wrap';
const PREVIEW_BG_KEY = 'office-svg-preview-bg';

export type SvgPreviewBackground = 'transparent' | 'white' | 'black';

export const SVG_PREVIEW_BACKGROUNDS: readonly SvgPreviewBackground[] = ['transparent', 'white', 'black'];

function isPreviewBackground(value: unknown): value is SvgPreviewBackground {
    return value === 'transparent' || value === 'white' || value === 'black';
}

export function loadSvgLineWrap(): boolean {
    const state = vscodeApi?.getState?.() as { svgLineWrap?: boolean } | undefined;
    if (state?.svgLineWrap !== undefined) {
        return state.svgLineWrap;
    }
    try {
        return localStorage.getItem(LINE_WRAP_KEY) === '1';
    } catch {
        return false;
    }
}

export function saveSvgLineWrap(enabled: boolean): void {
    try {
        localStorage.setItem(LINE_WRAP_KEY, enabled ? '1' : '0');
    } catch { }
    if (vscodeApi?.setState) {
        const prev = (vscodeApi.getState?.() ?? {}) as Record<string, unknown>;
        vscodeApi.setState({ ...prev, svgLineWrap: enabled });
    }
}

export function loadSvgPreviewBackground(): SvgPreviewBackground {
    const state = vscodeApi?.getState?.() as { svgPreviewBackground?: unknown } | undefined;
    if (isPreviewBackground(state?.svgPreviewBackground)) {
        return state.svgPreviewBackground;
    }
    try {
        const stored = localStorage.getItem(PREVIEW_BG_KEY);
        if (isPreviewBackground(stored)) {
            return stored;
        }
    } catch { }
    return 'transparent';
}

export function saveSvgPreviewBackground(background: SvgPreviewBackground): void {
    try {
        localStorage.setItem(PREVIEW_BG_KEY, background);
    } catch { }
    if (vscodeApi?.setState) {
        const prev = (vscodeApi.getState?.() ?? {}) as Record<string, unknown>;
        vscodeApi.setState({ ...prev, svgPreviewBackground: background });
    }
}
