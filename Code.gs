function onOpen()
{
  SpreadsheetApp.getUi().createMenu('Return to Dashboard')
      .addItem('Return to Dashboard', 'returnToDashboard')
      .addToUi();
}

function returnToDashboard()
{
  SpreadsheetApp.getActive().getSheetByName('Dashboard').activate()
}

function setSheetLinks()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheets = spreadsheet.getSheets();
  const dashboard = sheets.shift()
  const accountNums = dashboard.getSheetValues(4, 1, dashboard.getLastRow() - 3, 1)
  const ssURL = spreadsheet.getUrl();
  sheets.shift(); // Remove the customer list

  const sheetLinks = accountNums.map((acct, idx) => 
    [SpreadsheetApp.newRichTextValue().setText(acct[0]).setLinkUrl(ssURL + '#gid=' + 
      sheets[idx].getSheetId()).build()]
  )

  dashboard.getRange(4, 1, sheetLinks.length, 1).setRichTextValues(sheetLinks)
}

function removeSheets_()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheets = spreadsheet.getSheets();
  const sheetNames = sheets.map(s => s.getSheetName());

  for (var sheet = 0; sheet < sheetNames.length; sheet++)
  {
    if (sheetNames[sheet].split(' - ')[1])
      spreadsheet.deleteSheet(sheets[sheet])
  }
}

function createSheets_()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheets = spreadsheet.getSheets();
  const sheetNames = sheets.map(s => s.getSheetName());
  const customerList = sheets[sheetNames.indexOf('Customer List')];
  const templateSheet = sheets[sheetNames.indexOf('Template')]

  const customers = customerList.getSheetValues(3, 4, customerList.getLastRow() - 2, 1)

  for (var customer = 125; customer < customers.length; customer++)
    spreadsheet.insertSheet(customers[customer][0], 2 + customer, {template: templateSheet})
}

function createChart()
{
  const spreadsheet = SpreadsheetApp.getActive()
  const sheets = spreadsheet.getSheets();
  const sheetNames = sheets.map(sheet => sheet.getSheetName().split(' - '));
  const dashboard = sheets.shift();
  const initialYear = 2012;
  const numYears = 11;
  const row = dashboard.getActiveRange().getRow();
  const accountNumRng = dashboard.getRange(dashboard.getActiveRange().getRow(), 2);
  const accountNum = accountNumRng.getValue();
  const customerData = dashboard.getSheetValues(row, 3, 1, 13);
  const data = new Array(numYears).fill('').map((_,i) => [initialYear + i, customerData[0].pop()])
  const ssURL = spreadsheet.getUrl();
  var dataRng;

  sheetNames.shift() // Remove the Dashboard from the sheetNames array as well (since it was already removed from sheets array -- index needs to be identical)

  for (var s = 1; s < sheets.length; s++)
  {
    if (sheetNames[s][1] === accountNum)
    {
      // Set the Data
      sheets[s].getRange(1, 5, 1, 2).setValue('Chart Data').merge().setBorder(false, true, true, false, null, null)
      sheets[s].setColumnWidth(5, 75).setColumnWidth(6, 100).getRange(2, 5, 1, 2).setHorizontalAlignments([['center', 'right']]).setValues([['Year', 'Amount']])
      dataRng = sheets[s].getRange(3, 5, numYears, 2).setValues(data).clearContent().setBackground('white').setBorder(false, false, false, false, false, false).setFontWeight('normal')
        .setHorizontalAlignments(new Array(data.length).fill(['center', 'right'])).setNumberFormats(new Array(data.length).fill(['@', '$#,##0.00'])).setValues(data)

      const chart = dashboard.newChart()
        .asColumnChart()
        .addRange(dataRng)
        .setNumHeaders(0)
        .setXAxisTitle('Year')
        .setYAxisTitle('Sales Total')
        .setTransposeRowsAndColumns(false)
        .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
        .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
        .setOption('title', customerData[0][0])
        .setOption('subtitle', 'Total: $' + new Intl.NumberFormat().format(twoDecimals(customerData[0][1])))
        .setOption('isStacked', 'false')
        .setOption('bubble.stroke', '#000000')
        .setOption('textStyle.color', '#000000')
        .setOption('useFirstColumnAsDomain', true)
        .setOption('titleTextStyle.color', '#757575')
        .setOption('legend.textStyle.color', '#1a1a1a')
        .setOption('subtitleTextStyle.color', '#999999')
        .setOption('series', {0: {hasAnnotations: true, dataLabel: 'value'}})
        .setOption('trendlines', {0: {lineWidth: 4, type: 'linear', color: '#6aa84f'}})
        .setOption('hAxis', {textStyle: {color: '#000000'}, titleTextStyle: {color: '#000000'}})
        .setOption('annotations', {domain: {textStyle: {color: '#808080'}}, total: {textStyle : {color: '#808080'}}})
        .setOption('vAxes', {0: {textStyle: {color: '#000000'}, titleTextStyle: {color: '#000000'}, minorGridlines: {count: 2}}})
        .setPosition(1, 1, 0, 0)
        .build();

      dashboard.insertChart(chart);
      const steetId = spreadsheet.moveChartToObjectSheet(chart).activate().setName(sheetNames[s][0] + ' CHART - ' + sheetNames[s][1]).getSheetId();
      spreadsheet.moveActiveSheet(s + 3)
      accountNumRng.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(accountNum).setLinkUrl(
        ssURL + '#gid=' + steetId).build())
      
      break;
    }
  }
}

function createChartForSalesData()
{
  const numYears = 11;
  const spreadsheet = SpreadsheetApp.getActive()
  const dashboard = spreadsheet.getSheetByName('Dashboard')
  const dataRng = spreadsheet.getSheetByName('Sales Data').getRange(3, 1, numYears, 2)
  const grandTotal = dashboard.getSheetValues(3, 4, 1, 1)[0][0]
  const ssURL = spreadsheet.getUrl();

  const chart = dashboard.newChart()
    .asColumnChart()
    .addRange(dataRng)
    .setNumHeaders(0)
    .setXAxisTitle('Year')
    .setYAxisTitle('Sales Total')
    .setTransposeRowsAndColumns(false)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
    .setOption('title', 'Annual Lodge Sales Data')
    .setOption('subtitle', 'Total: $' + new Intl.NumberFormat().format(twoDecimals(grandTotal)))
    .setOption('isStacked', 'false')
    .setOption('bubble.stroke', '#000000')
    .setOption('textStyle.color', '#000000')
    .setOption('useFirstColumnAsDomain', true)
    .setOption('titleTextStyle.color', '#757575')
    .setOption('legend.textStyle.color', '#1a1a1a')
    .setOption('subtitleTextStyle.color', '#999999')
    .setOption('series', {0: {hasAnnotations: true, dataLabel: 'value'}})
    .setOption('trendlines', {0: {lineWidth: 4, type: 'linear', color: '#6aa84f'}})
    .setOption('hAxis', {textStyle: {color: '#000000'}, titleTextStyle: {color: '#000000'}})
    .setOption('annotations', {domain: {textStyle: {color: '#808080'}}, total: {textStyle : {color: '#808080'}}})
    .setOption('vAxes', {0: {textStyle: {color: '#000000'}, titleTextStyle: {color: '#000000'}, minorGridlines: {count: 2}}})
    .setPosition(1, 1, 0, 0)
    .build();

  dashboard.insertChart(chart);
  const steetId = spreadsheet.moveChartToObjectSheet(chart).activate().setName('ANNUAL LODGE SALES CHART').getSheetId();
  spreadsheet.moveActiveSheet(2)

  dashboard.getRange(1, 5).setRichTextValue(SpreadsheetApp.newRichTextValue().setText("Sale Totals").setLinkUrl(
    ssURL + '#gid=' + steetId).build())
}

function updateCustomersSalesData()
{
  const spreadsheet = SpreadsheetApp.getActive()
  const sheet = SpreadsheetApp.getActiveSheet()
  const accountNum = sheet.getSheetName().split(' - ')[1];
  const dashboard = SpreadsheetApp.getActive().getSheetByName('Dashboard');
  const accounts = dashboard.getSheetValues(4, 1, dashboard.getLastRow() - 3, 1).map(arr => arr[0].trim())
  const salesTotals = [new Array(11).fill('')]
  const numYears = 11;
  const currentYear = 2022; 
  var s, allCustomerAccounts, values, rowCounter = 0, year;
  var customerData = [[accountNum, , 'Total:', ''], ['Item Number', 'Item Description', 'Quantity', 'Amount']], 
       backgrounds = [['#c0c0c0', '#c0c0c0', '#c0c0c0', '#c0c0c0'], ['white', 'white', 'white', 'white']], 
       hAlignments = [['left', 'left', 'right', 'center'], ['left', 'left', 'right', 'right']]
       fontWeights = [['bold', 'bold', 'bold', 'bold'], ['bold', 'bold', 'bold', 'bold']],
     numberFormats = [['@', '@', '@', '$#,##0.00'], ['@', '@', '@', '@']];
         fontSizes = [[14, 14, 14, 14], [12, 12, 12, 12]]
        ranges = ['A1:D1'],
        formula = '=SUM(',
        customerName = null;

  for (var y = 0; y < numYears; y++)
  {
    year = (currentYear - y).toString();
    s = spreadsheet.getSheetByName(year);
    allCustomerAccounts = s.getSheetValues(22, 1, s.getLastRow() - 21, 1).map(arr => arr[0].trim())

    for (var i = allCustomerAccounts.length - 1; i >= 0; i--)
    {
      if (allCustomerAccounts[i] === accountNum)
      {
        if (y == 0) 
          customerData.push(['', '', 'Jan 01 ' + year, 'Oct 26 ' + year])
        else
          customerData.push(['', '', 'Jan 01 ' + year, 'Dec 31 ' + year])

        backgrounds.push(['white', 'white', 'white', 'white'])
        fontWeights.push(['bold', 'bold', 'bold', 'bold'])
        fontSizes.push([12, 12, 12, 12])
        hAlignments.push(['left', 'left', 'right', 'right'])
        numberFormats.push(['@', '@', '@', '@'])
        values = s.getSheetValues(i + 22, 12, rowCounter + 1, 17)

        if (customerName === null)
          customerName = s.getSheetValues(i + 22, 5, 1, 1)[0][0]

        for (j = 0; j < values.length; j++)
        {
          if (values[j][0] !== '')
          {
            customerData.push([values[j][0], values[j][6], values[j][14], values[j][16]])
            backgrounds.push(['white', 'white', 'white', 'white'])
            fontWeights.push(['normal', 'normal', 'normal', 'normal'])
            fontSizes.push([12, 12, 12, 12])
            hAlignments.push(['left', 'left', 'right', 'right'])
            numberFormats.push(['@', '@', '@', '$#,##0.00'])
          }
          else if (values[j][16] !== '')
          {
            customerData.push(['', '', values[j][14], values[j][16]])
            backgrounds.push(['white', 'white', '#c0c0c0', '#c0c0c0'])
            fontWeights.push(['normal', 'normal', 'bold', 'bold'])
            fontSizes.push([12, 12, 12, 12])
            hAlignments.push(['left', 'left', 'right', 'right'])
            ranges.push('C' + customerData.length + ':D' + customerData.length)
            formula += 'D' + customerData.length + ',';
            salesTotals[0][y] = values[j][16];
            numberFormats.push(['@', '@', '@', '$#,##0.00'])
            break;
          }
        }

        customerData.push(['', '', '', ''])
        backgrounds.push(['white', 'white', 'white', 'white'])
        fontWeights.push(['normal', 'normal', 'normal', 'normal'])
        fontSizes.push([12, 12, 12, 12])
        hAlignments.push(['left', 'left', 'right', 'right'])
        numberFormats.push(['@', '@', '@', '@'])
        rowCounter = 0;
        break;
      }
      else
        rowCounter++
    }
  }

  formula = formula.substring(0, formula.length - 1);
  formula += ')';

  customerData[0][1] = customerName;
  customerData[0][3] = formula;

  const numRows = customerData.length;
  const numCols = 4;
  const maxNumRows = sheet.getMaxRows();
  const maxNumCols = sheet.getMaxColumns();

  if (maxNumRows > numRows)
    sheet.deleteRows(numRows + 1, maxNumRows - numRows)
  if (maxNumCols > numCols)
    sheet.deleteColumns(5, maxNumCols - numCols)

  sheet.setRowHeight(1, 50).setColumnWidth(1, 225).setColumnWidth(2, 450).setColumnWidths(3, 2, 125).getRange(1, 1, numRows, numCols).breakApart().setWrap(false)
    .setBorder(false, false, false, false, false, false).setVerticalAlignment('middle').setHorizontalAlignments(hAlignments).setBackgrounds(backgrounds)
    .setFontWeights(fontWeights).setFontSizes(fontSizes).setNumberFormats(numberFormats).setValues(customerData)

  sheet.autoResizeRows(2, numRows - 1).setFrozenRows(2)

  sheet.getRangeList(ranges).setBorder(true, false, true, false, false, false);

  for (var u = 0; u < accounts.length; u++)
  {
    if (accounts[u] === accountNum)
    {
      dashboard.getRange(u + 4, 5, 1, salesTotals[0].length).setValues(salesTotals)
      break;
    }
  }
}

/**
 * This function take a number and rounds it to two decimals to make it suitable as a price.
 * 
 * @param {Number} num : The given number 
 * @return A number rounded to two decimals
 */
function twoDecimals(num)
{
  return Math.round((num + Number.EPSILON) * 100) / 100
}

function updatSalesDataWithPrompt()
{
  const ui = SpreadsheetApp.getUi()
  const response = ui.prompt('Provide Sheet Indecies:');

  if (response.getSelectedButton() == ui.Button.OK)
  {
    const index = response.getResponseText().split(' ')
    update_N_CustomersSalesData(index[0], index[1])
  } 
}

function update_N_CustomersSalesData(nnn, NNN)
{
  const spreadsheet = SpreadsheetApp.getActive()
  const sheets = spreadsheet.getSheets();
  const dashboard = sheets.shift();
  sheets.shift()
  const accounts = dashboard.getSheetValues(4, 1, dashboard.getLastRow() - 3, 1).map(arr => arr[0].trim())
  const numYears = 11;
  const currentYear = 2022; 
  var s, sheet, allCustomerAccounts, values, rowCounter = 0, year;

  for (var sh = nnn; sh < NNN; sh++)
  {
    var customerData = [['', , 'Total:', ''], ['Item Number', 'Item Description', 'Quantity', 'Amount']], 
         backgrounds = [['#c0c0c0', '#c0c0c0', '#c0c0c0', '#c0c0c0'], ['white', 'white', 'white', 'white']], 
         hAlignments = [['left', 'left', 'right', 'center'], ['left', 'left', 'right', 'right']]
         fontWeights = [['bold', 'bold', 'bold', 'bold'], ['bold', 'bold', 'bold', 'bold']],
       numberFormats = [['@', '@', '@', '$#,##0.00'], ['@', '@', '@', '@']];
           fontSizes = [[14, 14, 14, 14], [12, 12, 12, 12]]
              ranges = ['A1:D1'],
             formula = '=SUM(',
        customerName = null
         salesTotals = [new Array(numYears).fill('')];

    sheet = sheets[sh];
    var accountNum = sheet.getSheetName().split(' - ')[1];
    customerData[0][0] = accounts[sh]

    for (var y = 0; y < numYears; y++)
    {
      year = (currentYear - y).toString();
      s = spreadsheet.getSheetByName(year);
      allCustomerAccounts = s.getSheetValues(22, 1, s.getLastRow() - 21, 1).map(arr => arr[0].trim())

      for (var i = allCustomerAccounts.length - 1; i >= 0; i--)
      {
        if (allCustomerAccounts[i] === accountNum)
        {
          if (y == 0) 
            customerData.push(['', '', 'Jan 01 ' + year, 'Oct 26 ' + year])
          else
            customerData.push(['', '', 'Jan 01 ' + year, 'Dec 31 ' + year])

          backgrounds.push(['white', 'white', 'white', 'white'])
          fontWeights.push(['bold', 'bold', 'bold', 'bold'])
          fontSizes.push([12, 12, 12, 12])
          hAlignments.push(['left', 'left', 'right', 'right'])
          numberFormats.push(['@', '@', '@', '@'])
          values = s.getSheetValues(i + 22, 12, rowCounter + 1, 17)

          if (customerName === null)
            customerName = s.getSheetValues(i + 22, 5, 1, 1)[0][0]

          for (j = 0; j < values.length; j++)
          {
            if (values[j][0] !== '')
            {
              customerData.push([values[j][0], values[j][6], values[j][14], values[j][16]])
              backgrounds.push(['white', 'white', 'white', 'white'])
              fontWeights.push(['normal', 'normal', 'normal', 'normal'])
              fontSizes.push([12, 12, 12, 12])
              hAlignments.push(['left', 'left', 'right', 'right'])
              numberFormats.push(['@', '@', '@', '$#,##0.00'])
            }
            else if (values[j][16] !== '')
            {
              customerData.push(['', '', values[j][14], values[j][16]])
              backgrounds.push(['white', 'white', '#c0c0c0', '#c0c0c0'])
              fontWeights.push(['normal', 'normal', 'bold', 'bold'])
              fontSizes.push([12, 12, 12, 12])
              hAlignments.push(['left', 'left', 'right', 'right'])
              ranges.push('C' + customerData.length + ':D' + customerData.length)
              formula += 'D' + customerData.length + ',';
              salesTotals[0][y] = values[j][16];
              numberFormats.push(['@', '@', '@', '$#,##0.00'])
              break;
            }
          }

          customerData.push(['', '', '', ''])
          backgrounds.push(['white', 'white', 'white', 'white'])
          fontWeights.push(['normal', 'normal', 'normal', 'normal'])
          fontSizes.push([12, 12, 12, 12])
          hAlignments.push(['left', 'left', 'right', 'right'])
          numberFormats.push(['@', '@', '@', '@'])
          rowCounter = 0;
          break;
        }
        else
          rowCounter++
      }
    }

    formula = formula.substring(0, formula.length - 1);
    formula += ')';

    customerData[0][1] = customerName;
    customerData[0][3] = formula;

    const numRows = customerData.length;
    const numCols = 4;
    const maxNumRows = sheet.getMaxRows();
    const maxNumCols = sheet.getMaxColumns();

    if (maxNumRows > numRows)
      sheet.deleteRows(numRows + 1, maxNumRows - numRows)
    if (maxNumCols > numCols)
      sheet.deleteColumns(5, maxNumCols - numCols)

    sheet.setRowHeight(1, 50).setColumnWidth(1, 225).setColumnWidth(2, 450).setColumnWidths(3, 2, 125).getRange(1, 1, numRows, numCols).breakApart().setWrap(false)
      .setBorder(false, false, false, false, false, false).setVerticalAlignment('middle').setHorizontalAlignments(hAlignments).setBackgrounds(backgrounds)
      .setFontWeights(fontWeights).setFontSizes(fontSizes).setNumberFormats(numberFormats).setValues(customerData)

    sheet.autoResizeRows(2, numRows - 1).setFrozenRows(2)

    sheet.getRangeList(ranges).setBorder(true, false, true, false, false, false);

    for (var u = 0; u < accounts.length; u++)
    {
      if (accounts[u] === accountNum)
      {
        dashboard.getRange(u + 4, 5, 1, salesTotals[0].length).setValues(salesTotals)
        break;
      }
    }
  }
}

/**
 * This function takes the given string and makes sure that each word in the string has a capitalized 
 * first letter followed by lower case.
 * 
 * @param {String} str : The given string
 * @return {String} The output string with proper case
 * @author Jarren Ralf
 */
function toProper_(str)
{
  var numLetters;
  var words = str.toString().split(' ');

  for (var word = 0, string = ''; word < words.length; word++) 
  {
    numLetters = words[word].length;

    if (numLetters == 0) // The "word" is a blank string (a sentence contained 2 spaces)
      continue; // Skip this iterate
    else if (numLetters == 1) // Single character word
    {
      if (words[word][0] !== words[word][0].toUpperCase()) // If the single letter is not capitalized
        words[word] = words[word][0].toUpperCase(); // Then capitalize it
    }
    else
    {
      /* If the first letter is not upper case or the second letter is not lower case, then
       * capitalize the first letter and make the rest of the word lower case.
       */
      if (words[word][0] !== words[word][0].toUpperCase() || words[word][1] !== words[word][1].toLowerCase())
        words[word] = words[word][0].toUpperCase() + words[word].substring(1).toLowerCase();
    }

    string += words[word] + ' '; // Add a blank space at the end
  }

  string = string.slice(0, -1); // Remove the last space

  return string;
}

/**
* Gets the last row number based on a selected column range values
*
* @param {Object[][]} range Takes a 2d array of a single column's values
* @returns {Number} The last row number with a value. 
*/
function getLastRowSpecial(range)
{
  var rowNum = 0;
  var blank = false;
  
  for (var row = 0; row < range.length; row++)
  {
    if(range[row][0] === "" && !blank)
    {
      rowNum = row;
      blank = true;
    }
    else if (range[row][0] !== '')
      blank = false;
  }
  return (rowNum !== 0) ? rowNum : row;
}