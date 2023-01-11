const insertTableForNewMonth = function(monthName) {
  const ss = SpreadsheetApp.openById('###') // changed deliberately due to security
  const sheet = ss.getSheetByName('###'); // changed deliberately due to security
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
  const itemHeaderRange = sheet.getRange(2, startColumn, 1, 1)
  const amountHeaderRange = sheet.getRange(2, startColumn + 1, 1, 1);
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
  const openingBalanceHeadingRange = sheet.getRange(3, startColumn, 1, 1);
  const openingBalanceAmountRange = sheet.getRange(3, startColumn + 1, 1, 1);
  const currentBalanceHeadingRange = sheet.getRange(3, startColumn + 2, 1, 1);
  const currentBalanceAmountRange = sheet.getRange(3, startColumn + 3, 1, 1);
  const openingBalanceReferenceRange = sheet.getRange(3, openingBalanceAmountRange.getColumn() - 3);
  const openingBalanceAmount = openingBalanceReferenceRange.getValue();
  const amountsColumn = sheet.getRange(3, startColumn + 1)
  const currentBalanceAmount = sheet.getRange(1, startColumn + 1)

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

  // Row 4 - Salary and saved/(spent)

  // Row 5 - date of transfer
}

const getCurrentMonthName = function() {
  const date = new Date();
  const currentMonthName = date.toLocaleString('en-UK', {month: 'long'});
  return currentMonthName;
}