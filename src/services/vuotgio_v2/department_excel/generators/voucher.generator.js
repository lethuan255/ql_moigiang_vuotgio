/**
 * Voucher Generator — Renders the first two accounting voucher sheets in Type B exports:
 *   1. "Bảng kê chứng từ chi tiêu" (Mẫu số CQ05-HD)
 *   2. "Bảng kê chứng từ" (breakdown by training system funding sources)
 */

const WorkbookFactory   = require('../../shared_excel/core/workbook.factory');
const CellFormatter     = require('../../shared_excel/core/cell.formatter');
const PaymentCalculator = require('../data/calculator');

const SHEET_NAMES = {
    EXPENSE_VOUCHER: 'Bảng kê chứng từ chi tiêu',
    VOUCHER_SUMMARY: 'Bảng kê chứng từ',
};

const FONT_NAME = 'Times New Roman';
const MONEY_FMT = '#,##0.00';

const FUNDING_SYSTEM_COLUMNS = [
    { key: 'vn',     deptCol: 'AD', summaryCol: 'C', header: 'KINH PHÍ MẬT MÃ VN' },
    { key: 'lao',    deptCol: 'AE', summaryCol: 'D', header: 'KINH PHÍ LÀO' },
    { key: 'cuba',   deptCol: 'AF', summaryCol: 'E', header: 'KINH PHÍ CUBA' },
    { key: 'cpc',    deptCol: 'AG', summaryCol: 'F', header: 'KINH PHÍ CPC' },
    { key: 'dongHP', deptCol: 'AH', summaryCol: 'G', header: 'KINH PHÍ\nĐÓNG HỌC PHÍ' },
];

const _extractYearMeta = (namHoc) => {
    const raw = String(namHoc || '').trim();
    const compactNamHoc = raw.replace(/\s*-\s*/g, '-');
    const yearMatches = raw.match(/\d{4}/g);
    const endYear = yearMatches && yearMatches.length
        ? yearMatches[yearMatches.length - 1]
        : String(new Date().getFullYear());
    return { compactNamHoc, endYear };
};

const _quoteSheetName = (name) => `'${String(name || '').replace(/'/g, "''")}'`;

const _styleCell = (cell, {
    bold = false,
    italic = false,
    size = 11,
    hAlign = 'left',
    vAlign = 'middle',
    wrapText = false,
    border = false,
    numFmt,
    indent,
    fillArgb,
} = {}) => {
    cell.font = { name: FONT_NAME, size, bold, italic };
    cell.alignment = {
        horizontal: hAlign,
        vertical: vAlign,
        wrapText,
        ...(indent ? { indent } : {}),
    };
    if (numFmt) {
        cell.numFmt = numFmt;
    }
    if (fillArgb) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillArgb } };
    }
    if (border) {
        CellFormatter.applyBorder(cell);
    }
};

const _mergeAndStyle = (ws, r1, c1, r2, c2, value, styleOpts = {}) => {
    if (r1 !== r2 || c1 !== c2) {
        ws.mergeCells(r1, c1, r2, c2);
    }
    const cell = ws.getCell(r1, c1);
    cell.value = value;
    if (styleOpts.border) {
        for (let r = r1; r <= r2; r += 1) {
            for (let c = c1; c <= c2; c += 1) {
                CellFormatter.applyBorder(ws.getCell(r, c));
            }
        }
    }
    _styleCell(cell, styleOpts);
    return cell;
};

const _renderSignaturePair = (ws, startRow, leftSig, rightSig, blankRows = 5) => {
    const titleRow = startRow;
    const subRow = startRow + 1;
    const nameRow = subRow + blankRows + 1;

    _mergeAndStyle(ws, titleRow, leftSig.c1, titleRow, leftSig.c2, leftSig.title, {
        bold: true, size: 11, hAlign: 'center',
    });
    _mergeAndStyle(ws, titleRow, rightSig.c1, titleRow, rightSig.c2, rightSig.title, {
        bold: true, size: 11, hAlign: 'center',
    });
    ws.getRow(titleRow).height = 20;

    _mergeAndStyle(ws, subRow, leftSig.c1, subRow, leftSig.c2, leftSig.subtitle, {
        italic: true, size: 10.5, hAlign: 'center',
    });
    _mergeAndStyle(ws, subRow, rightSig.c1, subRow, rightSig.c2, rightSig.subtitle, {
        italic: true, size: 10.5, hAlign: 'center',
    });
    ws.getRow(subRow).height = 18;

    for (let i = 1; i <= blankRows; i += 1) {
        ws.getRow(subRow + i).height = 18;
    }

    _mergeAndStyle(ws, nameRow, leftSig.c1, nameRow, leftSig.c2, leftSig.name, {
        bold: true, size: 11, hAlign: 'center',
    });
    _mergeAndStyle(ws, nameRow, rightSig.c1, nameRow, rightSig.c2, rightSig.name, {
        bold: true, size: 11, hAlign: 'center',
    });
    ws.getRow(nameRow).height = 20;

    return nameRow;
};

const _resolveDeptMoneyTotals = (dept) => {
    if (dept.moneyTotals) {
        return dept.moneyTotals;
    }
    const totals = { vn: 0, lao: 0, cuba: 0, cpc: 0, dongHP: 0, total: 0 };
    (dept.summaries || []).forEach((sdo) => {
        const bd = sdo.breakdown || PaymentCalculator.computeSdoBreakdown(sdo.tableF, sdo.thanhToan, sdo.luong);
        totals.vn     += (bd.money?.vn     || 0);
        totals.lao    += (bd.money?.lao    || 0);
        totals.cuba   += (bd.money?.cuba   || 0);
        totals.cpc    += (bd.money?.cpc    || 0);
        totals.dongHP += (bd.money?.dongHP || 0);
        totals.total  += (bd.money?.total  || 0);
    });
    return {
        vn    : PaymentCalculator.excelNumber(totals.vn),
        lao   : PaymentCalculator.excelNumber(totals.lao),
        cuba  : PaymentCalculator.excelNumber(totals.cuba),
        cpc   : PaymentCalculator.excelNumber(totals.cpc),
        dongHP: PaymentCalculator.excelNumber(totals.dongHP),
        total : PaymentCalculator.excelNumber(totals.total),
    };
};

class VoucherGenerator {
    static get SHEET_NAMES() {
        return SHEET_NAMES;
    }

    /**
     * Sheet 1: "Bảng kê chứng từ chi tiêu" (Mẫu số CQ05-HD).
     *
     * @param {ExcelJS.Workbook} workbook
     * @param {object} opts
     * @param {Array}  opts.departmentList
     * @param {string} opts.namHoc
     * @param {string} [opts.summarySheetName='Bảng kê chứng từ']
     * @param {number} [opts.summaryStartRow=4]
     * @returns {ExcelJS.Worksheet}
     */
    static createExpenseVoucherSheet(workbook, {
        departmentList = [],
        namHoc,
        summarySheetName = SHEET_NAMES.VOUCHER_SUMMARY,
        summaryStartRow = 4,
    }) {
        const ws = WorkbookFactory.createWorksheet(workbook, SHEET_NAMES.EXPENSE_VOUCHER, {
            pageSetup: {
                orientation: 'portrait',
                paperSize: 9,
                fitToPage: true,
                fitToWidth: 1,
                fitToHeight: 1,
                margins: { left: 0.4, right: 0.4, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 },
            },
        });

        [6, 9, 8, 13, 42, 24].forEach((width, idx) => {
            ws.getColumn(idx + 1).width = width;
        });

        const { compactNamHoc, endYear } = _extractYearMeta(namHoc);
        const contentLabel = compactNamHoc
            ? `Thanh toán tiền vượt giờ năm học ${compactNamHoc}`
            : 'Thanh toán tiền vượt giờ';

        // Rows 1-8: Official CQ05-HD header
        ws.getRow(1).height = 10;

        _mergeAndStyle(ws, 2, 1, 2, 4, 'BAN CƠ YẾU CHÍNH PHỦ', { size: 11, hAlign: 'center' });
        _mergeAndStyle(ws, 2, 6, 2, 6, 'Mẫu số CQ05-HD', { bold: true, size: 11, hAlign: 'center' });
        ws.getRow(2).height = 20;

        _mergeAndStyle(ws, 3, 1, 3, 4, 'HỌC VIỆN KỸ THUẬT MẬT MÃ', { bold: true, size: 11, hAlign: 'center' });
        _mergeAndStyle(ws, 3, 6, 3, 6, 'Số: . . . . . . . . . . . . .', { size: 11, hAlign: 'center' });
        ws.getRow(3).height = 20;

        ws.getRow(4).height = 12;

        _mergeAndStyle(ws, 5, 1, 5, 6, 'BẢNG KÊ CHỨNG TỪ CHI TIÊU', {
            bold: true, size: 14, hAlign: 'center',
        });
        ws.getRow(5).height = 24;

        _mergeAndStyle(ws, 6, 1, 6, 6, `Tháng       năm ${endYear}`, {
            italic: true, size: 11, hAlign: 'center',
        });
        ws.getRow(6).height = 20;

        _mergeAndStyle(
            ws, 7, 1, 7, 6,
            'Kinh phí: Nguồn NSNN   Khoản: 085  Mục: 7750  Tiểu mục: 7799  Tiết mục: 03   Ngành: 00',
            { size: 11, hAlign: 'center' }
        );
        ws.getRow(7).height = 20;

        ws.getRow(8).height = 10;

        // Rows 9-11: Table header & column codes
        _mergeAndStyle(ws, 9, 1, 10, 1, 'STT', {
            bold: true, size: 11, hAlign: 'center', wrapText: true, border: true,
        });
        _mergeAndStyle(ws, 9, 2, 9, 4, 'Chứng từ, hoá đơn', {
            bold: true, size: 11, hAlign: 'center', wrapText: true, border: true,
        });
        _mergeAndStyle(ws, 10, 2, 10, 2, 'Loại', {
            bold: true, size: 11, hAlign: 'center', border: true,
        });
        _mergeAndStyle(ws, 10, 3, 10, 3, 'Số', {
            bold: true, size: 11, hAlign: 'center', border: true,
        });
        _mergeAndStyle(ws, 10, 4, 10, 4, 'Ngày', {
            bold: true, size: 11, hAlign: 'center', border: true,
        });
        _mergeAndStyle(ws, 9, 5, 10, 5, 'Nội dung chi tiêu', {
            bold: true, size: 11, hAlign: 'center', wrapText: true, border: true,
        });
        _mergeAndStyle(ws, 9, 6, 10, 6, 'Số tiền', {
            bold: true, size: 11, hAlign: 'center', wrapText: true, border: true,
        });
        ws.getRow(9).height = 22;
        ws.getRow(10).height = 22;

        ['A', 'B', 'C', 'D', 'E', 1].forEach((code, idx) => {
            const cell = ws.getCell(11, idx + 1);
            cell.value = code;
            _styleCell(cell, { bold: true, size: 10.5, hAlign: 'center', border: true });
        });
        ws.getRow(11).height = 20;

        // Rows 12+: Department voucher lines
        const dataStartRow = 12;
        let currentRow = dataStartRow;
        let grandTotal = 0;
        const summaryRef = summarySheetName ? _quoteSheetName(summarySheetName) : null;

        departmentList.forEach((dept, idx) => {
            const deptMoney = _resolveDeptMoneyTotals(dept);
            const rowAmount = PaymentCalculator.excelNumber(deptMoney.total);
            grandTotal += rowAmount;

            const row = ws.getRow(currentRow);
            row.getCell(1).value = idx + 1;
            _styleCell(row.getCell(1), { size: 11, hAlign: 'center', border: true });

            row.getCell(2).value = '';
            _styleCell(row.getCell(2), { size: 11, hAlign: 'center', border: true });

            row.getCell(3).value = String(idx + 1).padStart(2, '0');
            _styleCell(row.getCell(3), { size: 11, hAlign: 'center', numFmt: '@', border: true });

            row.getCell(4).value = '';
            _styleCell(row.getCell(4), { size: 11, hAlign: 'center', numFmt: '@', border: true });

            row.getCell(5).value = contentLabel;
            _styleCell(row.getCell(5), { size: 11, hAlign: 'left', indent: 1, wrapText: true, border: true });

            const summaryRow = summaryStartRow + idx;
            row.getCell(6).value = summaryRef
                ? { formula: `${summaryRef}!H${summaryRow}`, result: rowAmount }
                : rowAmount;
            _styleCell(row.getCell(6), {
                size: 11, hAlign: 'right', numFmt: MONEY_FMT, border: true,
            });

            row.height = 22;
            currentRow += 1;
        });

        const dataEndRow = currentRow - 1;
        const roundedGrandTotal = PaymentCalculator.excelNumber(grandTotal);

        // Total row
        const totalRow = currentRow;
        for (let col = 1; col <= 4; col += 1) {
            const cell = ws.getCell(totalRow, col);
            cell.value = '';
            _styleCell(cell, { size: 11, border: true });
        }
        const totalLabelCell = ws.getCell(totalRow, 5);
        totalLabelCell.value = 'Tổng cộng:';
        _styleCell(totalLabelCell, { bold: true, size: 11, hAlign: 'center', border: true });

        const totalAmountCell = ws.getCell(totalRow, 6);
        totalAmountCell.value = dataEndRow >= dataStartRow
            ? { formula: `SUM(F${dataStartRow}:F${dataEndRow})`, result: roundedGrandTotal }
            : 0;
        _styleCell(totalAmountCell, {
            bold: true, size: 11, hAlign: 'right', numFmt: MONEY_FMT, border: true,
        });
        ws.getRow(totalRow).height = 22;

        // Fixed amount & amount in words
        const fixedRow = totalRow + 1;
        _mergeAndStyle(ws, fixedRow, 1, fixedRow, 5, 'Ấn định số tiền chi bảng kê này là:', {
            bold: true, size: 11, hAlign: 'left', indent: 2,
        });
        const fixedAmountCell = ws.getCell(fixedRow, 6);
        fixedAmountCell.value = { formula: `F${totalRow}`, result: roundedGrandTotal };
        _styleCell(fixedAmountCell, { bold: true, size: 11, hAlign: 'right', numFmt: MONEY_FMT });
        ws.getRow(fixedRow).height = 22;

        const wordsRow = fixedRow + 1;
        const wordsText = PaymentCalculator.numberToWordsVi(Math.round(roundedGrandTotal));
        _mergeAndStyle(ws, wordsRow, 1, wordsRow, 6, wordsText, {
            bold: true, italic: true, size: 11, hAlign: 'left', wrapText: true,
        });
        ws.getRow(wordsRow).height = 24;

        // Date & first signature block
        const date1Row = wordsRow + 1;
        _mergeAndStyle(ws, date1Row, 5, date1Row, 6, `Ngày      tháng      năm ${endYear}`, {
            italic: true, size: 11, hAlign: 'center',
        });
        ws.getRow(date1Row).height = 20;

        const sig1EndRow = _renderSignaturePair(
            ws,
            date1Row + 1,
            { c1: 1, c2: 4, title: 'Phụ trách chi tiêu ngành', subtitle: '(Ký, họ tên)', name: 'Vương Thị Dung' },
            { c1: 5, c2: 6, title: 'Trưởng phòng Đào tạo',    subtitle: '(Ký, họ tên)', name: 'Đỗ Quang Trung' },
            5
        );

        // Accounting confirmation & second signature block
        const spacerRow = sig1EndRow + 1;
        ws.getRow(spacerRow).height = 12;

        const bookedRow = spacerRow + 1;
        _mergeAndStyle(ws, bookedRow, 1, bookedRow, 4, 'Đã ghi sổ cho thanh toán', {
            italic: true, size: 11, hAlign: 'center',
        });
        ws.getRow(bookedRow).height = 20;

        const date2Row = bookedRow + 1;
        _mergeAndStyle(ws, date2Row, 1, date2Row, 4, `Ngày ….. tháng … năm ${endYear}`, {
            italic: true, size: 11, hAlign: 'center',
        });
        ws.getRow(date2Row).height = 20;

        _renderSignaturePair(
            ws,
            date2Row + 1,
            { c1: 1, c2: 4, title: 'Cơ quan tài chính', subtitle: '(Ký, họ tên)',           name: 'Đặng Thị Hồng Nhung' },
            { c1: 5, c2: 6, title: 'Thủ trưởng đơn vị',  subtitle: '(Ký, họ tên, đóng dấu)', name: 'Hoàng Văn Thức' },
            5
        );

        return ws;
    }

    /**
     * Sheet 2: "Bảng kê chứng từ" (funding breakdown across all departments).
     *
     * @param {ExcelJS.Workbook} workbook
     * @param {object} opts
     * @param {Array} opts.departmentList - Each item may include `.sheetName`, `.totalRow`, `.moneyTotals`
     * @returns {ExcelJS.Worksheet}
     */
    static createVoucherSummarySheet(workbook, { departmentList = [] }) {
        const ws = WorkbookFactory.createWorksheet(workbook, SHEET_NAMES.VOUCHER_SUMMARY, {
            pageSetup: {
                orientation: 'landscape',
                paperSize: 9,
                fitToPage: true,
                fitToWidth: 1,
                fitToHeight: 1,
                margins: { left: 0.35, right: 0.35, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 },
            },
        });

        [6, 36, 18, 18, 18, 18, 20, 20].forEach((width, idx) => {
            ws.getColumn(idx + 1).width = width;
        });

        // Row 1: Title
        _mergeAndStyle(ws, 1, 1, 1, 8, 'BẢNG KÊ CHỨNG TỪ', {
            bold: true, size: 15, hAlign: 'center',
        });
        ws.getRow(1).height = 28;
        ws.getRow(2).height = 12;

        // Row 3: Column headers
        const headers = [
            'TT',
            'NỘI DUNG CHI',
            ...FUNDING_SYSTEM_COLUMNS.map((c) => c.header),
            'CỘNG NGANG',
        ];
        headers.forEach((label, idx) => {
            const cell = ws.getCell(3, idx + 1);
            cell.value = label;
            _styleCell(cell, {
                bold: true, size: 11, hAlign: 'center', wrapText: true, border: true,
            });
        });
        ws.getRow(3).height = 32;

        // Rows 4+: Department rows
        const dataStartRow = 4;
        let currentRow = dataStartRow;
        const grandTotals = { vn: 0, lao: 0, cuba: 0, cpc: 0, dongHP: 0, total: 0 };

        departmentList.forEach((dept, idx) => {
            const deptMoney = _resolveDeptMoneyTotals(dept);
            const row = ws.getRow(currentRow);

            row.getCell(1).value = idx + 1;
            _styleCell(row.getCell(1), { size: 11, hAlign: 'center', border: true });

            row.getCell(2).value = dept.khoa || dept.maKhoa || 'Khác';
            _styleCell(row.getCell(2), { size: 11, hAlign: 'left', indent: 1, wrapText: true, border: true });

            const canLinkDept = Boolean(dept.sheetName && dept.totalRow);
            const quotedDeptSheet = canLinkDept ? _quoteSheetName(dept.sheetName) : null;

            FUNDING_SYSTEM_COLUMNS.forEach((colDef, colOffset) => {
                const colIdx = 3 + colOffset;
                const val = PaymentCalculator.excelNumber(deptMoney[colDef.key]);
                grandTotals[colDef.key] += val;

                const cell = row.getCell(colIdx);
                cell.value = canLinkDept
                    ? { formula: `${quotedDeptSheet}!${colDef.deptCol}${dept.totalRow}`, result: val }
                    : val;
                _styleCell(cell, { size: 11, hAlign: 'right', numFmt: MONEY_FMT, border: true });
            });

            const rowTotal = PaymentCalculator.excelNumber(deptMoney.total);
            grandTotals.total += rowTotal;

            const rowSumCell = row.getCell(8);
            rowSumCell.value = { formula: `SUM(C${currentRow}:G${currentRow})`, result: rowTotal };
            _styleCell(rowSumCell, {
                bold: true, size: 11, hAlign: 'right', numFmt: MONEY_FMT, border: true,
            });

            row.height = 22;
            currentRow += 1;
        });

        const dataEndRow = currentRow - 1;
        const hasData = dataEndRow >= dataStartRow;

        const normalizedGrand = {
            vn    : PaymentCalculator.excelNumber(grandTotals.vn),
            lao   : PaymentCalculator.excelNumber(grandTotals.lao),
            cuba  : PaymentCalculator.excelNumber(grandTotals.cuba),
            cpc   : PaymentCalculator.excelNumber(grandTotals.cpc),
            dongHP: PaymentCalculator.excelNumber(grandTotals.dongHP),
            total : PaymentCalculator.excelNumber(grandTotals.total),
        };

        // Row: Cộng:
        const totalRow = currentRow;
        ws.getCell(totalRow, 1).value = '';
        _styleCell(ws.getCell(totalRow, 1), { size: 11, border: true });

        const sumLabelCell = ws.getCell(totalRow, 2);
        sumLabelCell.value = 'Cộng:';
        _styleCell(sumLabelCell, { bold: true, size: 11, hAlign: 'center', border: true });

        FUNDING_SYSTEM_COLUMNS.forEach((colDef, colOffset) => {
            const colIdx = 3 + colOffset;
            const cell = ws.getCell(totalRow, colIdx);
            cell.value = hasData
                ? { formula: `SUM(${colDef.summaryCol}${dataStartRow}:${colDef.summaryCol}${dataEndRow})`, result: normalizedGrand[colDef.key] }
                : 0;
            _styleCell(cell, { bold: true, size: 11, hAlign: 'right', numFmt: MONEY_FMT, border: true });
        });

        const totalHorizontalCell = ws.getCell(totalRow, 8);
        totalHorizontalCell.value = hasData
            ? { formula: `SUM(H${dataStartRow}:H${dataEndRow})`, result: normalizedGrand.total }
            : 0;
        _styleCell(totalHorizontalCell, {
            bold: true, size: 11, hAlign: 'right', numFmt: MONEY_FMT, border: true,
        });
        ws.getRow(totalRow).height = 22;

        // Row: Tổng cộng Mật mã (VN, Lào, Cuba, Campuchia) + ATTT:
        const grandRow = totalRow + 1;
        ws.getCell(grandRow, 1).value = '';
        _styleCell(ws.getCell(grandRow, 1), { size: 11, border: true });

        const grandLabelCell = ws.getCell(grandRow, 2);
        grandLabelCell.value = 'Tổng cộng Mật mã (VN, Lào, Cuba, Campuchia) + ATTT:';
        _styleCell(grandLabelCell, {
            bold: true, size: 11, hAlign: 'left', indent: 1, wrapText: true, border: true,
        });

        _mergeAndStyle(
            ws, grandRow, 3, grandRow, 8,
            { formula: `SUM(C${totalRow}:G${totalRow})`, result: normalizedGrand.total },
            { bold: true, size: 11, hAlign: 'center', numFmt: MONEY_FMT, border: true }
        );
        ws.getRow(grandRow).height = 30;

        // Row: Bằng chữ
        const wordsRow = grandRow + 1;
        const wordsText = `${PaymentCalculator.numberToWordsVi(Math.round(normalizedGrand.total))}.`;
        _mergeAndStyle(ws, wordsRow, 1, wordsRow, 8, wordsText, {
            bold: true, italic: true, size: 11, hAlign: 'left', wrapText: true,
        });
        ws.getRow(wordsRow).height = 26;

        // Signature block
        _renderSignaturePair(
            ws,
            wordsRow + 1,
            { c1: 1, c2: 4, title: 'Trưởng phòng Đào tạo',    subtitle: '(Ký, họ tên)', name: 'Đỗ Quang Trung' },
            { c1: 5, c2: 8, title: 'Người đề nghị thanh toán', subtitle: '(Ký, họ tên)', name: 'Vương Thị Dung' },
            5
        );

        return ws;
    }
}

module.exports = VoucherGenerator;
