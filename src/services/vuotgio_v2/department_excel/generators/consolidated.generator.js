/**
 * Consolidated Generator — Generates the full multi-department workbook (Type B).
 *
 * Sheet order:
 *   1. Sheet "Bảng kê chứng từ chi tiêu" — VoucherGenerator (Mẫu số CQ05-HD)
 *   2. Sheet "Bảng kê chứng từ"          — VoucherGenerator (tổng hợp kinh phí theo hệ)
 *   3. [N] Sheet per Khoa/Phòng          — DepartmentGenerator (36 columns per khoa)
 *      → isKhoa=0 units are merged into ONE sheet: "Ban giám đốc & các phòng"
 *   4. Sheet "TỔNG HỢP"                  — MasterSheetGenerator (mirrors tongHopGV.ejs)
 *   5. Sheet "Tiền chuyển khoản"         — PaymentGenerator (flat GV list)
 *
 * Luồng mới (post-snapshot):
 *   BẮT BUỘC đọc từ snapshot (vg_so_tiet_tong_hop). Năm chưa khóa → throw lỗi.
 */

const WorkbookFactory      = require('../../shared_excel/core/workbook.factory');
const DepartmentGenerator  = require('./department.generator');
const VoucherGenerator     = require('./voucher.generator');
const MasterSheetGenerator = require('./master.generator');
const PaymentGenerator     = require('./payment.generator');
const DataAggregator       = require('../data/aggregator');
const snapshotDataService  = require('../../snapshotData.service');

const _assignDepartmentSheetTargets = (departmentList) => {
    const usedNames = new Set([
        VoucherGenerator.SHEET_NAMES.EXPENSE_VOUCHER,
        VoucherGenerator.SHEET_NAMES.VOUCHER_SUMMARY,
        'TỔNG HỢP',
        WorkbookFactory.sanitizeWorksheetName('Tiền chuyển khoản', 'TienChuyenKhoan'),
    ]);

    departmentList.forEach((dept) => {
        const baseName = WorkbookFactory.sanitizeWorksheetName(dept.khoa || dept.maKhoa || 'Khác', 'PreviewKhoa');
        let candidate = baseName;
        let suffix = 2;
        while (usedNames.has(candidate)) {
            const tag = ` (${suffix})`;
            candidate = `${baseName.slice(0, 31 - tag.length)}${tag}`;
            suffix += 1;
        }
        usedNames.add(candidate);
        dept.sheetName = candidate;
        dept.totalRow = 6 + (dept.summaries ? dept.summaries.length : 0);
    });
};

class ConsolidatedGenerator {
    /**
     * Generate the full consolidated workbook for all departments.
     *
     * @param {string} namHoc
     * @returns {Promise<ExcelJS.Workbook>}
     */
    static async generateConsolidatedWorkbook(namHoc) {
        if (!namHoc) throw new Error('Thiếu thông tin Năm học');

        // ── 1. Fetch all SDOs from snapshot ─────────────────────────────────
        const allSummaries = await snapshotDataService.getSnapshotSDOList(namHoc, 'ALL');
        if (!allSummaries.length) throw new Error('Không có dữ liệu để xuất file');

        console.info('[ConsolidatedGenerator] start (snapshot)', {
            namHoc,
            totalSDOs: allSummaries.length,
        });

        // ── 2. Group by department & pre-resolve sheet targets ───────────────
        // - isKhoa=1 → one group per khoa
        // - isKhoa=0 → all merged into "Ban giám đốc & các phòng" (last in list)
        const departmentList = DataAggregator.groupByDepartment(allSummaries);
        _assignDepartmentSheetTargets(departmentList);

        // ── 3. Create workbook ───────────────────────────────────────────────
        const workbook = WorkbookFactory.createWorkbook({
            title  : `Tổng hợp vượt giờ ${namHoc}`,
            subject: `Báo cáo vượt giờ V2 ${namHoc}`,
            creator: 'VuotGioV2',
        });

        // ── 4. First 2 voucher sheets (CQ05-HD + Bảng kê chứng từ) ───────────
        VoucherGenerator.createExpenseVoucherSheet(workbook, { departmentList, namHoc });
        VoucherGenerator.createVoucherSummarySheet(workbook, { departmentList });

        // ── 5. Department sheets ─────────────────────────────────────────────
        for (const dept of departmentList) {
            const result = DepartmentGenerator.createDepartmentSheet(workbook, {
                khoa     : dept.khoa,
                maKhoa   : dept.maKhoa,
                sheetName: dept.sheetName,
                summaries: dept.summaries,
                namHoc,
                isExport : true,   // ← Excel formulas for dynamic recalculation
            });
            // Write back actual totals for use in master/payment sheets
            dept.totalThanhToan = result.totalThanhToan;
            dept.totalVuot      = result.totalVuot;
            dept.dataRowCount   = result.dataRowCount;
        }

        // ── 6. Master summary sheet (mirrors tongHopGV.ejs) ─────────────────
        MasterSheetGenerator.createMasterSheet(workbook, { departmentList, namHoc, isExport: true });

        // ── 7. Payment sheet (flat list of all lecturers, sorted by dept) ────
        PaymentGenerator.createPaymentSheet(workbook, {
            summaries: allSummaries,
            namHoc,
        });

        return workbook;
    }

    /**
     * Return structured preview data (no Excel generated).
     * Used by the web UI preview endpoint.
     *
     * @param {string} namHoc
     */
    static async getConsolidatedPreviewData(namHoc) {
        if (!namHoc) throw new Error('Thiếu thông tin Năm học');

        const allSummaries = await snapshotDataService.getSnapshotSDOList(namHoc, 'ALL');
        if (!allSummaries.length) throw new Error('Không có dữ liệu để xuất file');
        return DataAggregator.createConsolidatedData(namHoc, allSummaries);
    }
}

module.exports = ConsolidatedGenerator;