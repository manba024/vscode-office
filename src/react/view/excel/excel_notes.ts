import { XMLParser } from 'fast-xml-parser';
import type JSZip from 'jszip';
import type { CellNoteData, SheetData } from './x-spreadsheet/index';

type XmlRecord = Record<string, any>;

export type WorkbookNotes = Map<string, Map<string, CellNoteData>>;

const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '',
    removeNSPrefix: true,
    parseTagValue: false,
    trimValues: false,
});

const asArray = <T>(value: T | T[] | undefined): T[] => {
    if (value === undefined) return [];
    return Array.isArray(value) ? value : [value];
};

const readXmlText = (value: unknown): string => {
    if (value == null) return '';
    if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
        return String(value);
    }
    if (Array.isArray(value)) return value.map(readXmlText).join('');
    if (typeof value === 'object') {
        const node = value as XmlRecord;
        if (node['#text'] !== undefined) return readXmlText(node['#text']);
        if (node.t !== undefined) return readXmlText(node.t);
    }
    return '';
};

const readCommentText = (textNode: unknown): string => {
    if (!textNode || typeof textNode !== 'object') return readXmlText(textNode);
    const text = textNode as XmlRecord;
    if (text.t !== undefined) return readXmlText(text.t);
    const runs = asArray<XmlRecord>(text.r);
    if (runs.length > 0) return runs.map(run => readXmlText(run.t)).join('');
    return readXmlText(textNode);
};

const normalizePartPath = (sourcePart: string, target: string): string => {
    const normalizedTarget = target.replace(/\\/g, '/');
    const unresolved = normalizedTarget.startsWith('/')
        ? normalizedTarget.slice(1)
        : `${sourcePart.slice(0, sourcePart.lastIndexOf('/') + 1)}${normalizedTarget}`;
    const parts: string[] = [];
    unresolved.split('/').forEach((part) => {
        if (!part || part === '.') return;
        if (part === '..') {
            parts.pop();
            return;
        }
        parts.push(part);
    });
    return parts.join('/');
};

const relationshipPartPath = (partPath: string): string => {
    const slash = partPath.lastIndexOf('/');
    const directory = slash >= 0 ? partPath.slice(0, slash + 1) : '';
    const filename = slash >= 0 ? partPath.slice(slash + 1) : partPath;
    return `${directory}_rels/${filename}.rels`;
};

const parseRelationships = (xml: string): XmlRecord[] => {
    const root = parser.parse(xml) as XmlRecord;
    return asArray<XmlRecord>(root.Relationships?.Relationship);
};

const parseComments = (xml: string): Map<string, CellNoteData> => {
    const root = parser.parse(xml) as XmlRecord;
    const comments = asArray<XmlRecord>(root.comments?.commentList?.comment);
    const result = new Map<string, CellNoteData>();
    comments.forEach((comment) => {
        const address = typeof comment.ref === 'string' ? comment.ref.toUpperCase() : '';
        const text = readCommentText(comment.text);
        if (address && text.trim()) result.set(address, { text });
    });
    return result;
};

export function normalizeExcelJsNote(note: unknown): CellNoteData | undefined {
    if (typeof note === 'string') return note.trim() ? { text: note } : undefined;
    if (!note || typeof note !== 'object') return undefined;
    const texts = asArray<XmlRecord>((note as XmlRecord).texts);
    const text = texts.map(part => readXmlText(part.text)).join('');
    return text.trim() ? { text } : undefined;
}

/** Count notes already projected into spreadsheet data. */
export function countWorkbookNotes(sheets: readonly SheetData[]): number {
    return sheets.reduce((sheetTotal, sheet) => {
        const rows = Object.entries(sheet.rows ?? {});
        return sheetTotal + rows.reduce((rowTotal, [key, row]) => {
            if (key === 'len' || !row || typeof row !== 'object' || !('cells' in row)) return rowTotal;
            return rowTotal + Object.values(row.cells)
                .filter(cell => Boolean(cell?.note?.text))
                .length;
        }, 0);
    }, 0);
}

/** Read traditional Excel notes by following OOXML relationships. */
export async function readWorkbookNotes(zip: JSZip): Promise<WorkbookNotes> {
    const workbookPart = 'xl/workbook.xml';
    const workbookXml = await zip.file(workbookPart)?.async('string');
    const workbookRelsXml = await zip.file(relationshipPartPath(workbookPart))?.async('string');
    if (!workbookXml || !workbookRelsXml) return new Map();

    const workbook = parser.parse(workbookXml) as XmlRecord;
    const sheets = asArray<XmlRecord>(workbook.workbook?.sheets?.sheet);
    const relationshipById = new Map(
        parseRelationships(workbookRelsXml).map(rel => [String(rel.Id ?? ''), rel]),
    );
    const notesBySheet: WorkbookNotes = new Map();

    for (const sheet of sheets) {
        const sheetName = typeof sheet.name === 'string' ? sheet.name : '';
        const worksheetRelationship = relationshipById.get(String(sheet.id ?? ''));
        if (!sheetName || !worksheetRelationship?.Target) continue;
        if (!String(worksheetRelationship.Type ?? '').endsWith('/worksheet')) continue;

        const worksheetPart = normalizePartPath(workbookPart, String(worksheetRelationship.Target));
        const worksheetRelsXml = await zip.file(relationshipPartPath(worksheetPart))?.async('string');
        if (!worksheetRelsXml) continue;

        const commentsRelationship = parseRelationships(worksheetRelsXml)
            .find(rel => String(rel.Type ?? '').endsWith('/comments'));
        if (!commentsRelationship?.Target) continue;

        const commentsPart = normalizePartPath(worksheetPart, String(commentsRelationship.Target));
        const commentsXml = await zip.file(commentsPart)?.async('string');
        if (!commentsXml) continue;

        const notes = parseComments(commentsXml);
        if (notes.size > 0) notesBySheet.set(sheetName, notes);
    }

    return notesBySheet;
}
