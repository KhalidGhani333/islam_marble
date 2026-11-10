import ExcelJS from 'exceljs';
import type { WorkbookPass } from './sheetTypes';

function workbookAsBlob(buffer: ExcelJS.Buffer) {
  return new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

export async function downloadReportWorkbook(passes: WorkbookPass[]) {
  const report = new ExcelJS.Workbook();
  report.creator = 'Islam Marble';
  report.created = new Date();

  const registerSheet = report.addWorksheet('Gate Pass Register');
  registerSheet.columns = [
    { header: 'Gate Pass No.', key: 'gate_pass_number', width: 18 },
    { header: 'Date', key: 'date', width: 14 },
    { header: 'Person Name', key: 'person_name', width: 22 },
    { header: 'Mobile Number', key: 'mobile_number', width: 18 },
    { header: 'Company Name', key: 'company_name', width: 20 },
    { header: 'Total Items', key: 'total_items', width: 12 },
    { header: 'Total Pcs', key: 'total_pcs', width: 12 },
    { header: 'Total Feet', key: 'total_feet', width: 12 },
    { header: 'Total Amount', key: 'total_amount', width: 16 },
    { header: 'Description Summary', key: 'description_summary', width: 28 },
    { header: 'Taken By', key: 'taken_by', width: 18 },
    { header: 'Status', key: 'status', width: 16 },
    { header: 'Authorized By', key: 'authorized_by', width: 18 },
  ];
  const itemsSheet = report.addWorksheet('Gate Pass Items');
  itemsSheet.columns = [
    { header: 'Gate Pass No.', key: 'gate_pass_number', width: 18 },
    { header: 'S. No.', key: 'line_number', width: 10 },
    { header: 'Description', key: 'description', width: 30 },
    { header: 'Pieces', key: 'pieces', width: 14 },
    { header: 'Feet', key: 'feet', width: 14 },
    { header: 'Rate / Ft', key: 'rate', width: 14 },
    { header: 'Amount', key: 'amount', width: 16 },
  ];

  for (const pass of passes) {
    registerSheet.addRow({
      ...pass,
      description_summary: pass.items.map((item) => item.description).join('; '),
    });
    pass.items.forEach((item, index) => itemsSheet.addRow({
      gate_pass_number: pass.gate_pass_number,
      line_number: index + 1,
      description: item.description,
      pieces: item.pieces,
      feet: item.feet,
      rate: item.rate,
      amount: item.amount,
    }));
  }
  for (const sheet of [registerSheet, itemsSheet]) {
    sheet.views = [{ state: 'frozen', ySplit: 1 }];
    sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F2937' } };
  }

  const summarySheet = report.addWorksheet('Summary');
  summarySheet.columns = [{ header: 'Metric', width: 24 }, { header: 'Value', width: 18 }];
  summarySheet.addRows([
    ['Total Gate Passes', passes.length],
    ['Total Material Items', passes.reduce((total, pass) => total + pass.total_items, 0)],
    ['Total Pieces', passes.reduce((total, pass) => total + pass.total_pcs, 0)],
    ['Total Feet', passes.reduce((total, pass) => total + pass.total_feet, 0)],
    ['Total Amount', passes.reduce((total, pass) => total + pass.total_amount, 0)],
  ]);
  const summaryHeader = summarySheet.getRow(1);
  summaryHeader.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  summaryHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F2937' } };

  const url = URL.createObjectURL(workbookAsBlob(await report.xlsx.writeBuffer()));
  const link = document.createElement('a');
  link.href = url;
  link.download = 'islam-marble-gate-pass-report.xlsx';
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}
