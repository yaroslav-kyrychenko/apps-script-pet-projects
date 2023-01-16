const onOpen = function(e) {
  SpreadsheetApp.getUi()
    .createMenu('Custom functions')
    .addItem('Insert new savings/spending table', 'insertTableForNewMonth')
    .addToUi();
}

const insertTableForNewMonth = function (monthName) {
  // const ss = SpreadsheetApp.openById('######') // deliberately changed for security reasons
  // const sheet = ss.getSheetByName('###'); // deliberately changed for security reasons
  const startColumn = sheet.getLastColumn() + 2;
  const rules = sheet.getConditionalFormatRules();

  // Row 1 - month
  const monthNameRange = sheet.getRange(1, startColumn, 1, 4);
  setRangeFormat(monthNameRange, getMonthName(), 'center', 'bold', false, true);

  // Row 2 - headers
  const itemHeaderRange = sheet.getRange(2, startColumn,)
  const amountHeaderRange = sheet.getRange(2, startColumn + 1,);
  const statusHeaderRange = sheet.getRange(2, startColumn + 2, 1, 2)

  setRangeFormat(itemHeaderRange, 'Item', 'left', 'bold')
  setRangeFormat(amountHeaderRange, 'Amount', 'center', 'bold')
  setRangeFormat(statusHeaderRange, 'Status', 'center', 'bold')

  // Row 3 - Opening balance and current balance
  const openingBalanceHeadingRange = sheet.getRange(3, startColumn,);
  const openingBalanceAmountRange = sheet.getRange(3, startColumn + 1,);
  const currentBalanceHeadingRange = sheet.getRange(3, startColumn + 2,);
  const currentBalanceAmountRange = sheet.getRange(3, startColumn + 3,);
  const openingBalanceReferenceRange = sheet.getRange(3, openingBalanceAmountRange.getColumn() - 3);
  const openingBalanceAmount = openingBalanceReferenceRange.getValue();
  const openingBalanceAmountA1Notation = sheet.getRange(3, startColumn + 1).getA1Notation();
  let amountsColumnLetter;
  if (openingBalanceAmountA1Notation.length == 2) {
    amountsColumnLetter = openingBalanceAmountA1Notation[0];
  } 
  if (openingBalanceAmountA1Notation.length == 3) {
    amountsColumnLetter = [openingBalanceAmountA1Notation[0], openingBalanceAmountA1Notation[1]].join('');
  }
  const currentBalanceAmountFormula = `=SUM(${amountsColumnLetter}:${amountsColumnLetter})`;

  setRangeFormat(openingBalanceHeadingRange, 'Opening balance', 'left')
  setRangeFormat(openingBalanceAmountRange, openingBalanceAmount, 'center', 'normal', true)
  setRangeFormat(currentBalanceHeadingRange, 'Current balance', 'left', 'bold')
  setRangeFormat(currentBalanceAmountRange, currentBalanceAmountFormula, 'center', 'normal', true, false, true)

  // Row 4 - Salary and saved/(spent)
  const salaryHeadingRange = sheet.getRange(4, startColumn);
  const salaryAmountRange = sheet.getRange(4, startColumn + 1);
  const savedSpentHeadingRange = sheet.getRange(4, startColumn + 2);
  const savedSpentAmountRange = sheet.getRange(4, startColumn + 3);
  const savedSpentAmountFormula = `=${currentBalanceAmountRange.getA1Notation()}-${openingBalanceAmountRange.getA1Notation()}`;

  setRangeFormat(salaryHeadingRange, 'Salary', 'left')
  setRangeFormat(salaryAmountRange, '', 'center', 'normal', true)
  setRangeFormat(savedSpentHeadingRange, 'Saved/(spend)', 'left', 'bold')
  setRangeFormat(savedSpentAmountRange, savedSpentAmountFormula, 'center', 'normal', true, false, true);
  rules.push(...setConditionalFormatting(savedSpentAmountRange))
  // sheet.setConditionalFormatRules(setConditionalFormatting(savedSpentAmountRange));

  // Row 5 - date of transfer
  const dateOfTransferOfSalaryHeadingRange = sheet.getRange(5, startColumn, 1, 2);
  const dateOfTransferOfSalaryRange = sheet.getRange(5, startColumn + 2, 1, 2);

  setRangeFormat(dateOfTransferOfSalaryHeadingRange, 'Date of transfer of salary', 'left', 'normal', false, true)
  setRangeFormat(dateOfTransferOfSalaryRange, new Date(), 'center', 'normal', false, true);

  sheet.setConditionalFormatRules(rules);
  sheet.getRange(6, startColumn).activate();
}

const getMonthName = function () {
  const monthName = Browser.inputBox('Enter the month. Leave blank for the current month');
  if (monthName) return monthName;

  const date = new Date();
  const currentMonthName = date.toLocaleString('en-GB', { month: 'long' });
  if (!monthName) return currentMonthName;
}

const setRangeFormat = function (range, valueOrFormula, horizontalAlignment, fontWeight = 'normal', isNumber = false, isMerged = false, isFormula = false) {
  range
    .setHorizontalAlignment(horizontalAlignment)
    .setFontWeight(fontWeight)
    .setBorder(true, true, true, true, null, null)
  if (isNumber) range.setNumberFormat('0,000.00')
  if (isMerged) range.merge();
  isFormula ? range.setFormula(valueOrFormula) : range.setValue(valueOrFormula);
}

const setConditionalFormatting = function (range) {
  const cfRuleOverZero = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#34a853')
    .setFontColor('#ffffe3')
    .setRanges([range])
    .build()

  const cfRuleLessThanZero = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThanOrEqualTo(0)
    .setBackground('#ea4335')
    .setFontColor('#ffffe3')
    .setRanges([range])
    .build()

  let cfRules = [cfRuleOverZero, cfRuleLessThanZero]

  return cfRules;
}