const UPLOADHEADER = ["Reference #", "Receipt #", "G/L Account", "Amount", "Control #", "Description"]

enum COLUMN {
    RESPONSE = 15,
    DATE = 1,
    MERCHANT = 7,
    CHECK_NUMBER = 4,
    TOTAL_AMOUNT = 6,
    CONTROL = 21
}

enum ACCOUNTS {
    FIXED = 3225,
    VARIABLE = 3304,
    HOLD = 1000
}

interface UploadRow {
    reference: string;    // UTA091625(V,F,H)
    receipt: string;      // UTA091625(V,F,H)
    glAccount: number;    // 3225 || 3304
    amount: number;       // Total Trans Amount
    control: number;      // RO Num || Cust Num
    description: number;  // CHK #
}


function convertDate(excelDateValue: number) {
    const newDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
    const month = newDate.getMonth() + 1 > 9 ? newDate.getMonth() + 1 : `0${newDate.getMonth() + 1}`
    const day = newDate.getDate()
    const year = newDate.getFullYear().toString().slice(-2)
    return `${month}${day}${year}`
}


function getMerchantType(merchantCode: string) {
  let obj = { code: 'H', acct: ACCOUNTS.HOLD }
  if(merchantCode === '00') {
    obj = { code: 'V', acct: ACCOUNTS.VARIABLE }
  }
  if (merchantCode === '02') {
    obj = { code: 'F', acct: ACCOUNTS.FIXED }
  }
  return obj
}


function cleanRows(data: Array<string|number|boolean>[]) {
    const strippedData = data.filter(row => row[COLUMN.RESPONSE] != "DENIED")
    const newSheets = {}
    strippedData.forEach((row, index) => {
      if(index != 0) {
        const date = Number(row[COLUMN.DATE])
        const convertedDate = convertDate(date)
        const merch = getMerchantType(row[COLUMN.MERCHANT].toString().slice(-2))
        const refNum = `UTA${convertedDate}${merch.code}`

        const newRow = [refNum, refNum, merch.acct, row[COLUMN.TOTAL_AMOUNT], row[COLUMN.CONTROL], row[COLUMN.CHECK_NUMBER]]

        if (newSheets[refNum]) {
          newSheets[refNum].push(newRow)
        } else {
          newSheets[refNum] = [newRow]
        }
      }
    })
    return newSheets
}


function createSheet(wb: ExcelScript.Workbook, sheetName: string, data: Array<string | number>[]) {
    const sheet = wb.addWorksheet(sheetName)
    sheet.getRangeByIndexes(0, 0, 1, 6).setValues([UPLOADHEADER])
    sheet.getRangeByIndexes(1, 0, data.length, 6).setValues(data)
}


function main(workbook: ExcelScript.Workbook) {
    const sheets = workbook.getWorksheets()
    if(sheets.length > 1) {
        sheets.forEach((sheet, index) => {
            if(index > 0) sheet.delete()
        })
    }
    const reportSheet = sheets[0]
    const reportData = cleanRows(reportSheet.getUsedRange().getValues())
    const all_refs = Object.keys(reportData)
    all_refs.forEach(ref => createSheet(workbook, ref, reportData[ref]))

    const all_sheets = workbook.getWorksheets()
    all_sheets.forEach((sheet, index) => {
      if(index > 0) {
        sheet.getRange("1:1").getFormat().autofitColumns()
      }
    })
}
