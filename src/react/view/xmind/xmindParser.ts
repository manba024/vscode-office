import { convertXmindToMindElixir, importXMindFile } from '@mind-elixir/import-xmind';
import JSZip from 'jszip';
import type { MindElixirData, NodeObj } from 'mind-elixir';
import type { Sheet, Topic } from '@mind-elixir/import-xmind';

export interface XmindSheet {
    id: string;
    title: string;
    data: MindElixirData;
}

export interface XmindDocument {
    sheets: XmindSheet[];
    resolveImageUrl: (url: string) => string;
    export: (sheets: XmindSheet[]) => Promise<Uint8Array>;
    dispose: () => void;
}

const XAP_PREFIX = /^xap:(resources|attachments|resource)\//i;
const RESOURCE_DIRS = ['resources/', 'attachments/'];
const DEFAULT_SHEET_ID = 'sheet-1';
const DEFAULT_ROOT_TOPIC_ID = 'root-topic';

type MindArrow = NonNullable<MindElixirData['arrows']>[number];
type MindSummary = NonNullable<MindElixirData['summaries']>[number];

function createEmptyXmindSheet(): Sheet {
    return {
        id: DEFAULT_SHEET_ID,
        title: 'Sheet 1',
        rootTopic: {
            id: DEFAULT_ROOT_TOPIC_ID,
            title: '',
        },
        style: { id: '', type: '', properties: {} },
        topicPositioning: 'balanced',
        topicOverlapping: '',
        theme: { id: '', title: '' },
        relationships: [],
        legend: {
            visibility: 'visible',
            position: { x: 0, y: 0 },
            markers: {},
            groups: {},
        },
        settings: {
            'infoItems/infoItem': [],
            'tab-color': [],
        },
    };
}

function getMimeType(path: string): string {
    const lower = path.toLowerCase();
    if (lower.endsWith('.png')) {
        return 'image/png';
    }
    if (lower.endsWith('.jpg') || lower.endsWith('.jpeg')) {
        return 'image/jpeg';
    }
    if (lower.endsWith('.gif')) {
        return 'image/gif';
    }
    if (lower.endsWith('.webp')) {
        return 'image/webp';
    }
    if (lower.endsWith('.svg')) {
        return 'image/svg+xml';
    }
    if (lower.endsWith('.bmp')) {
        return 'image/bmp';
    }
    return 'application/octet-stream';
}

async function buildImageResourceMap(buffer: ArrayBuffer): Promise<{
    map: Map<string, string>;
    reverseMap: Map<string, string>;
    blobUrls: string[];
    zip: JSZip;
}> {
    const zip = await JSZip.loadAsync(buffer);
    const map = new Map<string, string>();
    const reverseMap = new Map<string, string>();
    const blobUrls: string[] = [];

    for (const [path, entry] of Object.entries(zip.files)) {
        if (entry.dir) {
            continue;
        }
        const normalizedPath = path.replace(/\\/g, '/');
        const lowerPath = normalizedPath.toLowerCase();
        if (!RESOURCE_DIRS.some(dir => lowerPath.startsWith(dir))) {
            continue;
        }

        const mime = getMimeType(normalizedPath);
        const data = await entry.async('uint8array');
        const blob = new Blob([data], { type: mime });
        const blobUrl = URL.createObjectURL(blob);
        blobUrls.push(blobUrl);

        const fileName = normalizedPath.split('/').pop() ?? normalizedPath;
        const zipPath = normalizedPath.replace(/^\/+/, '');
        const variants = new Set<string>([
            normalizedPath,
            zipPath,
            fileName,
            `xap:resources/${fileName}`,
            `xap:attachments/${fileName}`,
            `xap:resource/${fileName}`,
            `xap:resources/${zipPath}`,
            `xap:attachments/${zipPath}`,
            `xap:resource/${zipPath}`,
        ]);
        for (const key of variants) {
            map.set(key, blobUrl);
        }
        reverseMap.set(blobUrl, `xap:${zipPath}`);
    }

    return { map, reverseMap, blobUrls, zip };
}

function resolveXapResourceUrl(url: string, map: Map<string, string>): string {
    if (!url || !/^xap:/i.test(url)) {
        return url;
    }
    const direct = map.get(url) ?? map.get(url.toLowerCase());
    if (direct) {
        return direct;
    }

    const withoutScheme = url.replace(/^xap:/i, '');
    const fromScheme = map.get(withoutScheme) ?? map.get(withoutScheme.toLowerCase());
    if (fromScheme) {
        return fromScheme;
    }

    const zipPath = url.replace(XAP_PREFIX, '$1/');
    const fromZipPath = map.get(zipPath) ?? map.get(zipPath.toLowerCase());
    if (fromZipPath) {
        return fromZipPath;
    }

    const fileName = zipPath.split('/').pop();
    if (fileName) {
        const fromName = map.get(fileName) ?? map.get(fileName.toLowerCase());
        if (fromName) {
            return fromName;
        }
    }

    return url;
}

function patchNodeImages(node: NodeObj, resolve: (url: string) => string): void {
    if (node.image?.url) {
        node.image = {
            ...node.image,
            url: resolve(node.image.url),
        };
    }
    if (node.children) {
        for (const child of node.children) {
            patchNodeImages(child, resolve);
        }
    }
}

function cloneNodeWithoutParent(node: NodeObj): NodeObj {
    const { parent: _parent, children, ...rest } = node;
    return {
        ...rest,
        ...(children ? { children: children.map(cloneNodeWithoutParent) } : {}),
    };
}

function normalizeTags(tags: NodeObj['tags']): string | undefined {
    if (!tags?.length) {
        return undefined;
    }
    return tags.map(tag => typeof tag === 'string' ? tag : tag.text).filter(Boolean).join(',');
}

function mindNodeToXmindTopic(node: NodeObj, restoreImageUrl: (url: string) => string): Topic {
    const topic: Topic = {
        id: node.id,
        title: node.topic ?? 'Untitled',
    };
    if (node.expanded === false) {
        topic.branch = 'folded';
    }
    const labels = normalizeTags(node.tags);
    if (labels) {
        topic.labels = labels;
    }
    if (node.hyperLink) {
        topic.href = node.hyperLink;
    }
    if (node.note) {
        topic.notes = {
            plain: { content: node.note },
            html: { content: { paragraphs: [] } },
        };
    }
    if (node.image?.url) {
        topic.image = {
            src: restoreImageUrl(node.image.url),
            width: node.image.width,
            height: node.image.height,
            align: 'center',
        };
    }
    if (node.children?.length) {
        topic.children = {
            attached: node.children.map(child => mindNodeToXmindTopic(child, restoreImageUrl)),
        };
    }
    return topic;
}

function mindArrowsToXmindRelationships(arrows: MindArrow[] | undefined) {
    if (!arrows?.length) {
        return [];
    }
    return arrows.map(arrow => ({
        id: arrow.id,
        title: arrow.label || '',
        style: { id: '', type: '', properties: {} },
        class: '',
        end1Id: arrow.from,
        end2Id: arrow.to,
        controlPoints: {
            0: arrow.delta1 ?? { x: 50, y: 50 },
            1: arrow.delta2 ?? { x: 50, y: 50 },
        },
    }));
}

function collectSummaryTopics(topic: Topic, summaries: MindSummary[] | undefined): void {
    if (!summaries?.length) {
        return;
    }
    const topicSummaries = summaries.filter(summary => summary.parent === topic.id);
    if (topicSummaries.length) {
        topic.summaries = topicSummaries.map(summary => {
            const topicId = `${summary.id}-topic`;
            const summaryTopic: Topic = {
                id: topicId,
                title: summary.label || 'summary',
            };
            topic.children = topic.children ?? {};
            topic.children.summary = [...(topic.children.summary ?? []), summaryTopic];
            return {
                id: summary.id,
                style: { id: '', type: '', properties: {} },
                class: '',
                range: `(${summary.start},${summary.end})`,
                topicId,
            };
        });
    }
    topic.children?.attached?.forEach(child => collectSummaryTopics(child, summaries));
}

function mindDataToXmindSheet(original: Sheet, sheet: XmindSheet, restoreImageUrl: (url: string) => string): Sheet {
    const data = {
        ...sheet.data,
        nodeData: cloneNodeWithoutParent(sheet.data.nodeData),
    };
    const rootTopic = mindNodeToXmindTopic(data.nodeData, restoreImageUrl);
    collectSummaryTopics(rootTopic, data.summaries);
    return {
        ...original,
        id: sheet.id,
        title: sheet.title || data.nodeData.topic || original.title || 'Sheet',
        rootTopic,
        relationships: mindArrowsToXmindRelationships(data.arrows),
        topicPositioning: data.direction === 2 ? (original.topicPositioning || 'balanced') : original.topicPositioning,
    };
}

function createXmindDocument(
    sourceSheets: Sheet[],
    zip: JSZip,
    map: Map<string, string>,
    reverseMap: Map<string, string>,
    blobUrls: string[],
): XmindDocument {
    const result: XmindSheet[] = [];
    const resolveImageUrl = (url: string) => resolveXapResourceUrl(url, map);
    const restoreImageUrl = (url: string) => reverseMap.get(url) ?? url;
    for (const sheet of sourceSheets) {
        const data = convertXmindToMindElixir(sheet);
        patchNodeImages(data.nodeData, resolveImageUrl);
        result.push({
            id: sheet.id,
            title: sheet.title || 'Untitled',
            data,
        });
    }

    return {
        sheets: result,
        resolveImageUrl,
        export: async (nextSheets: XmindSheet[]) => {
            const sourceById = new Map(sourceSheets.map(sheet => [sheet.id, sheet]));
            const content = nextSheets.map(sheet => {
                const source = sourceById.get(sheet.id) ?? sourceSheets[0];
                return mindDataToXmindSheet(source, sheet, restoreImageUrl);
            });
            zip.file('content.json', JSON.stringify(content));
            zip.file('metadata.json', JSON.stringify({ creator: { name: 'Office Viewer' } }));
            zip.file('manifest.json', JSON.stringify({
                'file-entries': {
                    'content.json': {},
                    'metadata.json': {},
                },
            }));
            return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
        },
        dispose: () => {
            for (const url of blobUrls) {
                URL.revokeObjectURL(url);
            }
            map.clear();
            reverseMap.clear();
        },
    };
}

export async function parseXmind(buffer: ArrayBuffer, fileName = 'document.xmind'): Promise<XmindDocument> {
    if (buffer.byteLength === 0) {
        return createXmindDocument(
            [createEmptyXmindSheet()],
            new JSZip(),
            new Map<string, string>(),
            new Map<string, string>(),
            [],
        );
    }

    const { map, reverseMap, blobUrls, zip } = await buildImageResourceMap(buffer);
    const file = new File([buffer], fileName, { type: 'application/vnd.xmind.workbook' });
    const sourceSheets = await importXMindFile(file);
    return createXmindDocument(sourceSheets, zip, map, reverseMap, blobUrls);
}
