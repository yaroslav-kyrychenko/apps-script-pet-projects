const insertTableForNewMonth = function(monthName) {
  const ss = SpreadsheetApp.openById('###') // deliberately changed for security reasons
  const sheet = ss.getSheetByName('###'); // deliberately changed for security reasons
  const startColumn = sheet.getLastColumn() + 2;
  
  // Row 1 - month
  const monthNameRange = sheet.getRange(1, startColumn, 1, 4);
  monthNameRange
  .merge()
  .setValue(getCurrentMonthName())
  .setHorizontalAlignment('center')
  .setFontWeight('bold')
  .setBorder(true, true, true, true, null, null);
  
  // Row 2 - headers
  const itemHeaderRange = sheet.getRange(2, startColumn,)
  const amountHeaderRange = sheet.getRange(2, startColumn + 1,);
  const statusHeaderRange = sheet.getRange(2, startColumn + 2, 1, 2)
  
  itemHeaderRange
  .setValue('Item')
  .setHorizontalAlignment('left')
  .setFontWeight('bold')
  .setBorder(true, true, true, true, null, null);

  amountHeaderRange
  .setValue('Amount')
  .setHorizontalAlignment('center')
  .setFontWeight('bold')
  .setBorder(true, true, true, true, null, null);

  statusHeaderRange
  .merge()
  .setValue('Status')
  .setHorizontalAlignment('center')
  .setFontWeight('bold')
  .setBorder(true, true, true, true, null, null);

  // Row 3 - Opening balance and current balance
  const openingBalanceHeadingRange = sheet.getRange(3, startColumn,);
  const openingBalanceAmountRange = sheet.getRange(3, startColumn + 1,);
  const currentBalanceHeadingRange = sheet.getRange(3, startColumn + 2,);
  const currentBalanceAmountRange = sheet.getRange(3, startColumn + 3,);
  const openingBalanceReferenceRange = sheet.getRange(3, openingBalanceAmountRange.getColumn() - 3);
  const openingBalanceAmount = openingBalanceReferenceRange.getValue();
  const amountsColumnLetter = sheet.getRange(3, startColumn + 1).getA1Notation()[0];

  openingBalanceHeadingRange
  .setValue('Opening balance')
  .setHorizontalAlignment('left')
  .setBorder(true, true, true, true, null, null);

  openingBalanceAmountRange
  .setValue(openingBalanceAmount)
  .setHorizontalAlignment('center')
  .setNumberFormat('0,000.00')
  .setBorder(true, true, true, true, null, null);

  currentBalanceHeadingRange
  .setValue('Current balance')
  .setHorizontalAlignment('left')
  .setFontWeight('bold')
  .setBorder(true, true, true, true, null, null)

  currentBalanceAmountRange
  .setFormula(`=SUM(${amountsColumnLetter}:${amountsColumnLetter})`)
  .setHorizontalAlignment('center')
  .setNumberFormat('0,000.00')
  .setBorder(true, true, true, true, null, null)

  // Row 4 - Salary and saved/(spent)
  const salaryHeadingRange = sheet.getRange(4, startColumn);
  const salaryAmountRange = sheet.getRange(4, startColumn + 1);
  const savedSpentHeadingRange = sheet.getRange(4, startColumn + 2);
  const savedSpentAmountRange = sheet.getRange(4, startColumn + 3);

  salaryHeadingRange
  .setValue('Salary')
  .setHorizontalAlignment('left')
  .setBorder(true, true, true, true, null, null)

  salaryAmountRange
  .setHorizontalAlignment('center')
  .setNumberFormat('0,000.00')
  .setBorder(true, true, true, true, null, null)
  
  savedSpentHeadingRange
  .setValue('Saved/(spent)')
  .setHorizontalAlignment('left')
  .setFontWeight('bold')
  .setBorder(true, true, true, true, null, null)

  savedSpentAmountRange
  .setFormula(`=${currentBalanceAmountRange.getA1Notation()}-${openingBalanceAmountRange.getA1Notation()}`)
  .setHorizontalAlignment('center')
  .setNumberFormat('0,000.00')
  .setBorder(true, true, true, true, null, null)

  // Row 5 - date of transfer
  const dateOfTransferOfSalaryHeadingRange = sheet.getRange(5, startColumn);
  const dateOfTransferOfSalaryRange = sheet.getRange(5, startColumn + 1);
  
  dateOfTransferOfSalaryHeadingRange
  .setValue('Date of transfer of salary')
  .setHorizontalAlignment('left')
  .setBorder(true, true, true, true, null, null)

  dateOfTransferOfSalaryRange
  .setValue(getDateOfTransferOfSalary())
  .setHorizontalAlignment('center')
  .setBorder(true, true, true, true, null, null);

  Logger.log(typeof getDateOfTransferOfSalary())
  Logger.log(typeof sheet.getRange('B5'))
}

const getCurrentMonthName = function() {
  const date = new Date();
  const currentMonthName = date.toLocaleString('en-GB', {month: 'long'});
  return currentMonthName;
}

const getDateOfTransferOfSalary = function() {
  const date = new Date();
  const dateOfTransferOfSalary = date.toLocaleString('en-GB', {
    year: 'numeric',
    month: 'numeric',
    day: 'numeric'
  })
  return dateOfTransferOfSalary;
}