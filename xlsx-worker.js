/* global XLSX */
const XLSX_URLS = [
    'vendor/xlsx.full.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
];

let workbook = null;
let activeSheetName = null;
let worksheet = null;
let sheetInfo = null;

function loadXlsx() {
    if (self.XLSX) return;
    for (const url of XLSX_URLS) {
        try {
            importScripts(url);
            if (self.XLSX) {
                return;
            }
        } catch (error) {
            // Try next URL
        }
    }
    throw new Error('XLSX library failed to load in worker.');
}

function getSheetInfo(worksheetRef) {
    const range = XLSX.utils.decode_range(worksheetRef['!ref'] || 'A1');
    const rowCount = Math.max(0, range.e.r - range.s.r + 1);
    const colCount = Math.max(0, range.e.c - range.s.c + 1);
    const headerRows = XLSX.utils.sheet_to_json(worksheetRef, {
        header: 1,
        range: { s: { r: 0, c: 0 }, e: { r: 0, c: Math.max(colCount - 1, 0) } },
        defval: '',
        blankrows: false
    });
    const headers = headerRows[0] || [];
    const normalized = [];
    for (let i = 0; i < colCount; i++) {
        normalized[i] = headers[i] !== undefined && headers[i] !== '' ? headers[i] : `(Column ${indexToColumnLetter(i)})`;
    }
    return { rowCount, colCount, headers: normalized };
}

function indexToColumnLetter(index) {
    let letter = '';
    let i = index;
    while (i >= 0) {
        letter = String.fromCharCode((i % 26) + 65) + letter;
        i = Math.floor(i / 26) - 1;
    }
    return letter;
}

function setActiveSheet(name) {
    if (!workbook) {
        throw new Error('Workbook not initialized');
    }
    const sheetNames = workbook.SheetNames || [];
    activeSheetName = name || sheetNames[0];
    worksheet = workbook.Sheets[activeSheetName];
    if (!worksheet) {
        throw new Error(`Sheet "${activeSheetName}" not found`);
    }
    sheetInfo = getSheetInfo(worksheet);
    return sheetInfo;
}

self.onmessage = (event) => {
    const { id, type, buffer, sheetName, startRow, endRow, maxCols } = event.data || {};
    try {
        if (type === 'init') {
            loadXlsx();
            workbook = XLSX.read(buffer, { type: 'array' });
            const sheetNames = workbook.SheetNames || [];
            if (sheetNames.length > 0) {
                setActiveSheet(sheetName || sheetNames[0]);
            }
            self.postMessage({ id, ok: true, data: { sheetNames } });
            return;
        }

        if (!workbook || !worksheet) {
            throw new Error('Workbook not loaded');
        }

        if (type === 'selectSheet') {
            const info = setActiveSheet(sheetName);
            self.postMessage({ id, ok: true, data: info });
            return;
        }

        if (type === 'getInfo') {
            const info = sheetInfo || getSheetInfo(worksheet);
            self.postMessage({ id, ok: true, data: info });
            return;
        }

        if (type === 'getRows') {
            const safeMaxCols = Math.max(maxCols || 0, 0);
            if (safeMaxCols === 0) {
                self.postMessage({ id, ok: true, data: [] });
                return;
            }
            const range = {
                s: { r: startRow, c: 0 },
                e: { r: endRow, c: Math.max(safeMaxCols - 1, 0) }
            };
            const rows = XLSX.utils.sheet_to_json(worksheet, {
                header: 1,
                range,
                defval: '',
                blankrows: true
            });
            self.postMessage({ id, ok: true, data: rows });
            return;
        }

        throw new Error(`Unknown worker message type: ${type}`);
    } catch (error) {
        self.postMessage({ id, ok: false, error: error.message || String(error) });
    }
};
