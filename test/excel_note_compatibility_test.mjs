import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';

import ExcelJS from '@cweijan/exceljs';
import JSZip from 'jszip';
import { readExcel } from '../src/react/view/excel/excel_reader.ts';
import { countWorkbookNotes } from '../src/react/view/excel/excel_notes.ts';
import { renderCell } from '../src/react/view/excel/x-spreadsheet/component/table.js';

const OPENPYXL_FIXTURE = path.join(process.cwd(), 'test/fixtures/excel/openpyxl-note.xlsx');
const EXTERNAL_SAMPLE = process.env.EXCEL_NOTE_SAMPLE;

function noteText(note) {
    if (typeof note === 'string') return note;
    return note?.texts?.map((part) => part.text ?? '').join('') ?? '';
}

async function loadWorkbook(buffer) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    return workbook;
}

function toArrayBuffer(buffer) {
    if (buffer instanceof ArrayBuffer) return buffer;
    return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

async function addLargeSparseValidation(buffer) {
    const zip = await JSZip.loadAsync(buffer);
    const worksheetPath = 'xl/worksheets/sheet1.xml';
    const worksheetXml = await zip.file(worksheetPath)?.async('string');
    assert.ok(worksheetXml);
    assert.equal(worksheetXml.includes('<dataValidations'), false);

    const validationXml = [
        '<dataValidations count="1">',
        '<dataValidation type="list" allowBlank="1" sqref="A1:A1000 A1002:A2000">',
        '<formula1>"one,two"</formula1>',
        '</dataValidation>',
        '</dataValidations>',
    ].join('');
    zip.file(worksheetPath, worksheetXml.replace('</worksheet>', `${validationXml}</worksheet>`));
    return zip.generateAsync({ type: 'nodebuffer' });
}

test('reads traditional notes written by OpenPyXL', async () => {
    const workbook = await loadWorkbook(await readFile(OPENPYXL_FIXTURE));
    const worksheet = workbook.getWorksheet('Notes');

    assert.ok(worksheet);
    assert.equal(noteText(worksheet.getCell('A1').note), 'OpenPyXL note');
});

test('projects traditional notes into read-only spreadsheet cell data', async () => {
    const data = await readExcel(toArrayBuffer(await readFile(OPENPYXL_FIXTURE)));
    const worksheet = data.sheets.find((sheet) => sheet.name === 'Notes');

    assert.ok(worksheet);
    assert.deepEqual(worksheet.rows?.[0]?.cells[0]?.note, { text: 'OpenPyXL note' });
    assert.deepEqual(worksheet.rows?.[2]?.cells[2]?.note, { text: 'Note on empty cell' });
    assert.equal(countWorkbookNotes(data.sheets), 2);
});

test('ignores large sparse data validations in read-only preview', async () => {
    const source = await readFile(OPENPYXL_FIXTURE);
    const buffer = await addLargeSparseValidation(source);
    const data = await readExcel(toArrayBuffer(buffer));
    const worksheet = data.sheets.find((sheet) => sheet.name === 'Notes');

    assert.ok(worksheet);
    assert.equal(worksheet.validations?.length ?? 0, 0);
    assert.deepEqual(worksheet.rows?.[0]?.cells[0]?.note, { text: 'OpenPyXL note' });
});

test('renders a note indicator for cells with read-only note data', () => {
    const previousDocument = globalThis.document;
    const previousGetComputedStyle = globalThis.getComputedStyle;
    globalThis.document = {
        querySelector: () => null,
        documentElement: {},
    };
    globalThis.getComputedStyle = () => ({ getPropertyValue: () => '' });

    try {
        let noteIndicators = 0;
        const cell = { text: 'value', note: { text: 'Read-only note' } };
        const draw = {
            rect: (_box, callback) => callback(),
            text: () => {},
            note: () => { noteIndicators += 1; },
            error: () => {},
            frozen: () => {},
            strokeBorders: () => {},
        };
        const data = {
            rows: { isHide: () => false, getCell: () => cell },
            cols: { isHide: () => false },
            getCell: () => cell,
            canEditCell: () => true,
            getCellStyleOrDefault: () => ({ font: { name: 'Arial', size: 11 } }),
            defaultStyle: () => ({ font: { name: 'Arial', size: 11 } }),
            cellRect: () => ({ left: 0, top: 0, width: 100, height: 30 }),
            getZoomScale: () => 1,
            getHyperlink: () => undefined,
            getValidationError: () => undefined,
            sortedRowMap: new Map(),
            merges: null,
            settings: { evalPaused: true },
        };

        renderCell(draw, data, 0, 0);
        assert.equal(noteIndicators, 1);
    } finally {
        globalThis.document = previousDocument;
        globalThis.getComputedStyle = previousGetComputedStyle;
    }
});

test('keeps note text when an optional VML drawing part is unavailable', async () => {
    const zip = await JSZip.loadAsync(await readFile(OPENPYXL_FIXTURE));
    zip.remove('xl/drawings/commentsDrawing1.vml');
    const buffer = await zip.generateAsync({ type: 'nodebuffer' });

    const workbook = await loadWorkbook(buffer);
    const worksheet = workbook.getWorksheet('Notes');
    assert.ok(worksheet);
    assert.equal(noteText(worksheet.getCell('A1').note), 'OpenPyXL note');

    const spreadsheet = await readExcel(toArrayBuffer(buffer));
    assert.equal(countWorkbookNotes(spreadsheet.sheets), 2);
});

test('continues to read the traditional-note layout written by ExcelJS', async () => {
    const source = new ExcelJS.Workbook();
    const sourceSheet = source.addWorksheet('Canonical');
    sourceSheet.getCell('B2').value = 'value';
    sourceSheet.getCell('B2').note = 'Canonical note';
    sourceSheet.getCell('C3').note = {
        texts: [
            { text: 'Rich ' },
            { font: { bold: true }, text: 'note' },
        ],
    };

    const buffer = await source.xlsx.writeBuffer();
    const workbook = await loadWorkbook(buffer);
    assert.equal(noteText(workbook.getWorksheet('Canonical').getCell('B2').note), 'Canonical note');

    const spreadsheet = await readExcel(toArrayBuffer(buffer));
    assert.deepEqual(spreadsheet.sheets[0].rows?.[1]?.cells[1]?.note, { text: 'Canonical note' });
    assert.deepEqual(spreadsheet.sheets[0].rows?.[2]?.cells[2]?.note, { text: 'Rich note' });
});

test('continues to read workbooks without notes', async () => {
    const source = new ExcelJS.Workbook();
    source.addWorksheet('Plain').getCell('A1').value = 'plain';

    const workbook = await loadWorkbook(await source.xlsx.writeBuffer());
    assert.equal(workbook.getWorksheet('Plain').getCell('A1').value, 'plain');
});

if (EXTERNAL_SAMPLE) {
    test('loads the external workbook and projects every traditional note', async () => {
        const buffer = await readFile(EXTERNAL_SAMPLE);
        const workbook = await loadWorkbook(buffer);
        let noteCount = 0;

        workbook.eachSheet((worksheet) => {
            worksheet.eachRow({ includeEmpty: true }, (row) => {
                row.eachCell({ includeEmpty: true }, (cell) => {
                    if (noteText(cell.note)) noteCount += 1;
                });
            });
        });

        assert.ok(workbook.worksheets.length > 0);
        assert.ok(noteCount > 0);

        const spreadsheet = await readExcel(toArrayBuffer(buffer));
        assert.equal(countWorkbookNotes(spreadsheet.sheets), noteCount);
    });
}
