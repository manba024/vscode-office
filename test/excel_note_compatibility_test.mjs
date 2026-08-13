import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';

import ExcelJS from '@cweijan/exceljs';
import JSZip from 'jszip';

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

test('reads traditional notes written by OpenPyXL', async () => {
    const workbook = await loadWorkbook(await readFile(OPENPYXL_FIXTURE));
    const worksheet = workbook.getWorksheet('Notes');

    assert.ok(worksheet);
    assert.equal(noteText(worksheet.getCell('A1').note), 'OpenPyXL note');
});

test('keeps note text when an optional VML drawing part is unavailable', async () => {
    const zip = await JSZip.loadAsync(await readFile(OPENPYXL_FIXTURE));
    zip.remove('xl/drawings/commentsDrawing1.vml');

    const workbook = await loadWorkbook(await zip.generateAsync({ type: 'nodebuffer' }));
    const worksheet = workbook.getWorksheet('Notes');

    assert.ok(worksheet);
    assert.equal(noteText(worksheet.getCell('A1').note), 'OpenPyXL note');
});

test('continues to read the traditional-note layout written by ExcelJS', async () => {
    const source = new ExcelJS.Workbook();
    const sourceSheet = source.addWorksheet('Canonical');
    sourceSheet.getCell('B2').value = 'value';
    sourceSheet.getCell('B2').note = 'Canonical note';

    const workbook = await loadWorkbook(await source.xlsx.writeBuffer());
    assert.equal(noteText(workbook.getWorksheet('Canonical').getCell('B2').note), 'Canonical note');
});

test('continues to read workbooks without notes', async () => {
    const source = new ExcelJS.Workbook();
    source.addWorksheet('Plain').getCell('A1').value = 'plain';

    const workbook = await loadWorkbook(await source.xlsx.writeBuffer());
    assert.equal(workbook.getWorksheet('Plain').getCell('A1').value, 'plain');
});

if (EXTERNAL_SAMPLE) {
    test('loads the external workbook and exposes at least one traditional note', async () => {
        const workbook = await loadWorkbook(await readFile(EXTERNAL_SAMPLE));
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
    });
}
