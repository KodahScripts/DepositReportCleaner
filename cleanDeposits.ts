class Amount {
  public total: number;
  constructor(public auth: number, public sett: number, public fee: number) {
    this.total = sett - fee;
  }
}

class Row {
  public state: string;
  public brand: string;
  public amount: Amount;
  public invoice: string;
  public user: string;
  constructor(protected row: Array<string | number | boolean>) {
    this.state = String(row[1]);
    this.brand = String(row[2]);
    this.invoice = String(row[6]).length > 1 ? String(row[6]) : String(row[7]);

    const authAmount = this.getAmount(String(row[3]));
    const settAmount = this.getAmount(String(row[4]));
    const feeAmount = this.getAmount(String(row[5]));
    this.amount = new Amount(authAmount, settAmount, feeAmount);

    this.user = String(row[8]);
  }

  getAmount(cellValue: string): number {
    if (cellValue === "") return 0;
    if (!cellValue.includes("$")) return 0;

    const amt = cellValue.split("$")[1];
    if (!amt || amt.length === 0) return 0;

    return Number(amt.replace(/,/g, ""));
  }
}

function main(workbook: ExcelScript.Workbook) {
  if(workbook.getWorksheets().length > 1) {
    workbook.getWorksheets().forEach((sht, idx) => {
      if(idx !== 0) sht.delete();
    });
  }

  const selectedSheet = workbook.getActiveWorksheet();
  let reportData = selectedSheet.getUsedRange().getValues();
  const rowLength = reportData[0].filter(cell => cell !== "");
  reportData = rowLength.length > 9 ? removeCols(selectedSheet, reportData) : reportData;

  const data = reportData.map((row, index) => {
    return new Row(row);
  });

  const approvedReceipts = data.filter(d => d.state === "APPROVAL");
  const settledReceipts = data.filter(d => d.state === "SETTLED");

  const approvedAX = approvedReceipts.filter(rec => rec.brand === "AMEX");
  const approvedOther = approvedReceipts.filter(rec => rec.brand !== "AMEX");
  const settledAX = settledReceipts.filter(rec => rec.brand === "AMEX");
  const settledOther = settledReceipts.filter(rec => rec.brand !== "AMEX");

  if (approvedAX.length > 0) addSheet(workbook, "Approved AX", approvedAX);
  if (approvedOther.length > 0) addSheet(workbook, "Approved Other", approvedOther);
  if (settledAX.length > 0) addSheet(workbook, "Settled AX", settledAX);
  if (settledOther.length > 0) addSheet(workbook, "Settled Other", settledOther);
}

function addSheet(workbook: ExcelScript.Workbook, sheetName: string, data: Row[]) {
  const sheet = workbook.addWorksheet(sheetName);
  const totalRow = data.length + 2;
  sheet.getRange("A1:F1").setValues([["Auth Amount", "Settlement Amount", "Cardholder Surcharge", "Total", "Invoice Number", "User"]]);
  data.forEach((d, i) => {
    const row = i + 2;
    const invoice = getInvoiceNumber(d.invoice);
    sheet.getRange(`A${row}:F${row}`).setValues([[d.amount.auth, d.amount.sett, d.amount.fee, `=IFERROR(B${row}-C${row},B${row})`, invoice, d.user]]);
  });

  const range = sheet.getUsedRange();
  const table = sheet.addTable(range, true);
  const tableLen = table.getRowCount() + 1;
  table.addRow(null, [
    `=SUM(A2:A${tableLen})`,
    `=SUM(B2:B${tableLen})`,
    `=SUM(C2:C${tableLen})`,
    '', '', ''
  ]);
  table.getHeaderRowRange().getFormat().autofitColumns();
}

function getInvoiceNumber(cellValue: string): string {
  if(cellValue.length < 6) return cellValue;
  if(cellValue.slice(-6).split('').filter(val => Number(val) !== 0).length === 2) return cellValue.slice(0,6);
  return cellValue.slice(-6);
}

function removeCols(sheet: ExcelScript.Worksheet, reportData: Array<string | number | boolean>[]): Array<string|number|boolean>[] {
  sheet.getRange("A:D").delete(ExcelScript.DeleteShiftDirection.left);
  sheet.getRange("B:C").delete(ExcelScript.DeleteShiftDirection.left);
  sheet.getRange("D:H").delete(ExcelScript.DeleteShiftDirection.left);
  sheet.getRange("I:J").delete(ExcelScript.DeleteShiftDirection.left);
  sheet.getRange("J:K").delete(ExcelScript.DeleteShiftDirection.left);
  return sheet.getUsedRange(true).getValues();
}
