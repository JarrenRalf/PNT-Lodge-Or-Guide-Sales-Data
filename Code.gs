/**
 * This function processes the imported data.
 * 
 * @param {Event Object} e : The event object from an installed onChange trigger.
 */
function onChange(e)
{
  try
  {
    processImportedData(e)
  }
  catch (error)
  {
    Logger.log(error['stack'])
    Browser.msgBox(error['stack'])
  }
}

/**
 * This function places a menu at the 
 */
function onOpen()
{
  SpreadsheetApp.getUi().createMenu('Return to Dashboard')
    .addItem('Return to Dashboard', 'returnToDashboard')
    .addToUi();
}

/**
 * This function adds a new customer to the customer list and to the Dashboard. It then creates a template for their data sheet and chart.
 * 
 * @author Jarren Ralf
 */
function addNewCustomer()
{
  const ui = SpreadsheetApp.getUi()
  const response1 = ui.prompt('What is the customer number?')

  if (response1.getSelectedButton() === ui.Button.OK)
  {
    const response2 = ui.prompt('What is the customer name?')

    if (response2.getSelectedButton() === ui.Button.OK)
    {
      const response3 = ui.prompt('What is the abbreviated customer name?')

      if (response3.getSelectedButton() === ui.Button.OK)
      {
        const customerNumber = response1.getResponseText().toUpperCase()
          const customerName = response2.getResponseText().toUpperCase()
             const sheetName = response3.getResponseText().toUpperCase() + ' - ' + customerNumber

        if (isNotBlank(customerNumber) && isNotBlank(customerName) && isNotBlank(sheetName))
        {
          const response4 = ui.alert('You entered the following information:\nCustomer #: \t' + customerNumber + '\nCustomer Name: \t' + customerName + '\nSheet Name: \t' + sheetName + '\n\nDoes this look correct?',ui.ButtonSet.YES_NO)

          if (response4 === ui.Button.YES)
          {
            const spreadsheet = SpreadsheetApp.getActive();
            const dashboard = spreadsheet.getSheetByName('Dashboard')
            const customerSheet = spreadsheet.getSheetByName('Customer List')
            var numRows = customerSheet.getLastRow() - 2;
            const customerNumbers = customerSheet.getSheetValues(3, 1, numRows, 1).flat()

            if (customerNumbers.includes(customerNumber))
              ui.alert('Customer is already in the list.')
            else
            {
              const customerSheetList = SpreadsheetApp.openById('1xKw4GAtNbAsTEodCDmCMbPCbXUlK9OHv0rt5gYzqx9c').getSheetByName('Customer List');
              const customerList = customerSheetList.getSheetValues(2, 1, customerSheetList.getLastRow() - 1, 2);
              const numCustomers = customerList.push([customerNumber, customerName])
              customerSheetList.getRange(2, 1, numCustomers, 2).setValues(customerList.sort((a, b) => (a[0] > b[0]) ? 1 : (a[0] < b[0]) ? -1 : 0)) // Sort the customers by CUST #
              
              numRows++;
              const range = customerSheet.appendRow([customerNumber, customerName, sheetName]).getRange(3, 1, numRows, 3)
              const values = range.getValues().sort((a, b) => (a[1] > b[1]) ? 1 : (a[1] < b[1]) ? -1 : 0) // Sort the customers alphabetically by name
              range.setValues(values)

              // Once list is sorted, find new customer, then identify the customer that proceeds the new one, we will use this insert the customer sheets into the correct location
              const newCustIndex = values.findIndex(custNum => custNum[0] === customerNumber)

              if (newCustIndex !== 0) // If the new customer is not alphabetically first in the list
              {
                const previousCustomerNum = values[newCustIndex - 1][0]; 
                const sheetNames = spreadsheet.getSheets().map(sht => sht.getSheetName().split(' - '))

                // Figure out what the index should be for the customer data sheet
                for (var i = 4; i < sheetNames.length; i++)
                  if (sheetNames[i][1] === previousCustomerNum)
                    break;
              }
              else // The new customer is first
                var i = 2;

              const customerDataSheet = spreadsheet.insertSheet(sheetName, (i + 2), {template: spreadsheet.getSheetByName('Template')}).showSheet()
              const id_chart = createChart_NewCustomer(customerName, sheetName, customerDataSheet, spreadsheet) // Store ID so that we can hyperlink to the chart from Dashboard
              const lastRow = dashboard.getLastRow() + 1;
              const numCols = dashboard.getLastColumn();
              const sheetLinks = dashboard.getRange(4, 1, lastRow - 4, 2).getRichTextValues()
              const formulas_CustomerTotals = dashboard.getRange(4, 4, lastRow - 4).getFormulas()
              dashboard.appendRow(['', '', customerName, ...new Array(numCols - 3).fill('')])
              const dashboardValues = dashboard.getSheetValues(4, 1, lastRow - 3, numCols).sort((a, b) => (a[2] > b[2]) ? 1 : (a[2] < b[2]) ? -1 : 0) // Sort customer alphabetically

              formulas_CustomerTotals.splice(newCustIndex, 0, 
                ['=SUM(E' + (newCustIndex + 4) + ':' + dashboard.getRange(1, numCols).getA1Notation()[0] + (newCustIndex + 4) + ')']
              )

              for (var i = newCustIndex; i < formulas_CustomerTotals.length; i++) // The formulas need to be adjusted slightl, increase by one below the new customer
                formulas_CustomerTotals[i][0] = formulas_CustomerTotals[i][0].replaceAll(/\d+/g, i + 4)

              sheetLinks.splice(newCustIndex, 0, // Set hyperlinks for the new sheet
                [SpreadsheetApp.newRichTextValue().setText(customerNumber).setLinkUrl('#gid=' + customerDataSheet.getSheetId()).build(), 
                 SpreadsheetApp.newRichTextValue().setText(customerNumber).setLinkUrl('#gid=' + id_chart).build()]
              )

              const formulaRange_YearlyTotals = dashboard.getRange(4, 1, lastRow - 3, numCols).setValues(dashboardValues) // Set the customer names and sales values
                .offset(0, 0, lastRow - 3, 2).setRichTextValues(sheetLinks)        // Set the hyperlinked sheet links
                .offset(0, 3, lastRow - 3, 1).setFormulas(formulas_CustomerTotals) // Set the customer totals formulas
                .offset(-1, 0, 1, numCols - 3); // The range of the yearly sales totals; Their formulas need tp be updated because we have an extra row

              const formulas_YearlyTotals = [formulaRange_YearlyTotals.getFormulas()[0].map(formula => formula.toString().substring(0, 9) + lastRow.toString() + ')')]
              formulaRange_YearlyTotals.setFormulas(formulas_YearlyTotals).offset(1, -3, lastRow - 3, numCols) // Set new formulas because we have a new customer and they need to extend one more row
              dashboard.getRange(newCustIndex + 4, 1, 1, 2).activate() // Move to the sheet links
            }
          }
        }
        else
          ui.alert('Atleast one of your typed responses was blank.\n\nPlease redo the process.')
      }
    }
  }
}

/**
 * This function configures the yearly customer item data into the format that is desired for the spreadsheet to function optimally
 * 
 * @param {Object[][]}      values         : The values of the data that were just imported into the spreadsheet
 * @param {String}         fileName        : The name of the new sheet (which will also happen to be the xlxs file name)
 * @param {Boolean} doesPreviousSheetExist : Whether the previous sheet with the same name exists or not
 * @param {Spreadsheet}   spreadsheet      : The active spreadsheet
 * @author Jarren Ralf
 */
function configureYearlyCustomerItemData(values, fileName, doesPreviousSheetExist, spreadsheet)
{
  const currentYear = new Date().getFullYear();
  const customerSheet = spreadsheet.getSheetByName('Customer List');
  const accounts = customerSheet.getSheetValues(3, 1, customerSheet.getLastRow() - 2, 1).map(v => v[0].toString().trim())
  values.shift()
  values.pop() // Remove the final row which contains descriptive stats
  const preData = values.filter(d => accounts.includes(d[0].toString().trim()));
  const [data, ranges] = reformatData_YearlyCustomerItemData(preData)
  const yearRange = new Array(currentYear - 2012 + 1).fill('').map((_, y) => (currentYear - y).toString()).reverse()
  var year = yearRange.find(p => p == fileName) // The year that the data is representing

  if (year == null)
  {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt('Enter the year:')

    if (response.getSelectedButton() === ui.Button.OK)
    {
      year = response.getResponseText(); // Text response is assumed to be the year

      if (yearRange.includes(year))
      {
        const numCols = 6;
        const sheets = spreadsheet.getSheets();
        const previousSheet = sheets.find(sheet => sheet.getSheetName() == year)
        var indexAdjustment = 2010

        if (previousSheet != null)
        {
          indexAdjustment--;
          spreadsheet.deleteSheet(previousSheet)
        }

        SpreadsheetApp.flush();
        const newSheet = spreadsheet.insertSheet(year, sheets.length + indexAdjustment - year)
          .setColumnWidth(1, 66).setColumnWidth(2, 300).setColumnWidth(3, 150).setColumnWidth(4, 300).setColumnWidths(5, 2, 75);
        SpreadsheetApp.flush();
        const lastRow = data.unshift(['Customer', 'Customer Name', 'Item Number', 'Item Description', 'Quantity', 'Amount']);
        newSheet.deleteColumns(7, 20)
        newSheet.setTabColor('#a64d79').setFrozenRows(1)
        newSheet.protect()
        newSheet.getRange(1, 1, 1, numCols).setFontSize(11).setFontWeight('bold').setBackground('#c0c0c0')
          .offset(0, 0, lastRow, numCols).setHorizontalAlignments(new Array(lastRow).fill(['left', 'left', 'left', 'left', 'right', 'right'])).setNumberFormat('@').setValues(data)
        newSheet.getRangeList(ranges).setBorder(true, false, true, false, false, false).setBackground('#c0c0c0').setFontWeight('bold')

        const dashboard = spreadsheet.getSheetByName('Dashboard')

        if (currentYear > Number(dashboard.getRange('E2').getValue()))
        {
          const dashboard_lastRow = dashboard.getLastRow();
          dashboard.insertColumnBefore(5).getRange(2, 5, 2, 1).setValues([[currentYear], ['=SUM(E4:E' + dashboard_lastRow + ')']])
          const grandTotalRange = dashboard.getRange(4, 4, dashboard_lastRow - 3)
          grandTotalRange.setFormulas(grandTotalRange.getFormulas().map(formula => [formula[0].replace('F', 'E')]))
        }

        updateAllCustomersSalesData(spreadsheet)
      }
      else
      {
        ui.alert('Invalid Input', 'Please import your data again and enter a 4 digit year in the range of [2012, ' + currentYear + '].',)
        return;
      }
    }
    else
    {
      spreadsheet.toast('Data import proccess has been aborted.', '', 60)
      return;
    }
  }
  else
  {
    const numCols = 6;
    const sheets = spreadsheet.getSheets();
    const previousSheet = sheets.find(sheet => sheet.getSheetName() == year)
    var indexAdjustment = 2011

    if (doesPreviousSheetExist)
    {
      indexAdjustment--;
      spreadsheet.deleteSheet(previousSheet)
    }
    
    SpreadsheetApp.flush();
    const newSheet = spreadsheet.insertSheet(year, sheets.length + indexAdjustment - year)
      .setColumnWidth(1, 66).setColumnWidth(2, 300).setColumnWidth(3, 150).setColumnWidth(4, 300).setColumnWidths(5, 2, 75);
    SpreadsheetApp.flush();
    const lastRow = data.unshift(['Customer', 'Customer Name', 'Item Number', 'Item Description', 'Quantity', 'Amount']);
    newSheet.deleteColumns(7, 20)
    newSheet.setTabColor('#a64d79').setFrozenRows(1)
    newSheet.protect()
    newSheet.getRange(1, 1, 1, numCols).setFontSize(11).setFontWeight('bold').setBackground('#c0c0c0')
      .offset(0, 0, lastRow, numCols).setHorizontalAlignments(new Array(lastRow).fill(['left', 'left', 'left', 'left', 'right', 'right'])).setNumberFormat('@').setValues(data)
    newSheet.getRangeList(ranges).setBorder(true, false, true, false, false, false).setBackground('#c0c0c0').setFontWeight('bold')

    const dashboard = spreadsheet.getSheetByName('Dashboard')

    if (currentYear > Number(dashboard.getRange('E2').getValue()))
    {
      const dashboard_lastRow = dashboard.getLastRow();
      dashboard.insertColumnBefore(5).getRange(2, 5, 2, 1).setValues([[currentYear], ['=SUM(E4:E' + dashboard_lastRow + ')']])
      const grandTotalRange = dashboard.getRange(4, 4, dashboard_lastRow - 3)
      grandTotalRange.setFormulas(grandTotalRange.getFormulas().map(formula => [formula[0].replace('F', 'E')]))
    }

    updateAllCustomersSalesData(spreadsheet)
  }
}

/**
 * This function creates an embedded column chart on a new sheet based on the active spreadsheet that the user is on.
 * 
 * @author Jarren Ralf
 */
function createChart()
{
  const spreadsheet = SpreadsheetApp.getActive()
  const activeSheet = SpreadsheetApp.getActiveSheet();
  const numYears = new Date().getFullYear() - 2012 + 1
  const dataRng = activeSheet.getRange(3, 5, numYears, 2);
  const customerValues = activeSheet.getRange(1, 2, 1, 3).getDisplayValues()[0]

  const chart = activeSheet.newChart()
    .asColumnChart()
    .addRange(dataRng)
    .setNumHeaders(0)
    .setXAxisTitle('Year')
    .setYAxisTitle('Sales Total')
    .setTransposeRowsAndColumns(false)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
    .setOption('title', customerValues[0])
    .setOption('subtitle', 'Total: ' + customerValues[2])
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

  activeSheet.insertChart(chart);
  const sheetNameSplit = activeSheet.getSheetName().split(' - ');
  const sheetName_CHART = sheetNameSplit[0] + ' CHART - ' + sheetNameSplit[1];
  spreadsheet.deleteSheet(spreadsheet.getSheetByName(sheetName_CHART))
  spreadsheet.moveChartToObjectSheet(chart).activate().setName(sheetNameSplit[0] + ' CHART - ' + sheetNameSplit[1])
}

/**
 * This function creates a chart sheet for the new customer that is being created by the user.
 * 
 * @param    {String}  customerName    : The name of the customer
 * @param    {String}    sheetName     : The name of the customer's data sheet
 * @param    {Sheet} customerDataSheet : The sheet containing the customer's data
 * @param {Spreadsheet} spreadsheet    : The active spreadsheet
 * @return {Number} The id of the sheet object that is created for the chart
 * @author Jarren Ralf
 */
function createChart_NewCustomer(customerName, sheetName, customerDataSheet, spreadsheet)
{
  const currentYear = new Date().getFullYear();
  const chartData = new Array(currentYear - 2012 + 1).fill('').map((_, y) => [(currentYear - y).toString(), '']).reverse()
  const numRows = chartData.length;
  const sheetName_Split = sheetName.split(' - ')
  
  const chartDataRng = customerDataSheet.setTabColor('#38761d').getRange(3, 5, numRows, 2).setBackground('white').setBorder(false, false, false, false, false, false)
    .setFontWeight('normal').setHorizontalAlignments(new Array(numRows).fill(['center', 'right'])).setNumberFormats(new Array(numRows).fill(['@', '$#,##0.00'])).setValues(chartData)
  customerDataSheet.setColumnWidth(5, 75).setColumnWidth(6, 100)
    .getRange(1, 1, 1, 5).setValues([[sheetName_Split[1], customerName, 'Total:', '=SUM(' + customerDataSheet.getRange(3, 6, numRows).getA1Notation() + ')', 'Chart Data']])
    .offset(0, 4, 1, 2).merge().setBorder(false, true, true, false, null, null)
    .offset(1, 0, 1, 2).setHorizontalAlignments([['center', 'right']]).setValues([['Year', 'Amount']])

  const chart = customerDataSheet.newChart()
    .asColumnChart()
    .addRange(chartDataRng)
    .setNumHeaders(0)
    .setXAxisTitle('Year')
    .setYAxisTitle('Sales Total')
    .setTransposeRowsAndColumns(false)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
    .setOption('title', customerName)
    .setOption('subtitle', 'Total: $' + new Intl.NumberFormat().format(twoDecimals(0)))
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

  customerDataSheet.insertChart(chart);
  
  return spreadsheet.moveChartToObjectSheet(chart).activate().setName(sheetName_Split[0] + ' CHART - ' + sheetName_Split[1]).setTabColor('#f1c232').getSheetId();
}

/**
 * This function takes the active selection on the Dashboard and deletes the customer from the customer list on both the current spreadsheet and the 
 * Lodge, Charter, & Guide spreadsheet. It also deletes the customer data sheet and chart sheet.
 * 
 * @author Jarren Ralf
 */
function deleteSelectedCustomers()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const customerListSheet = spreadsheet.getSheetByName('Customer List');
  const customerList = customerListSheet.getSheetValues(3, 1, customerListSheet.getLastRow() - 2, 3);
  const customerListSheet_LodgeAndCharterSS = SpreadsheetApp.openById('1xKw4GAtNbAsTEodCDmCMbPCbXUlK9OHv0rt5gYzqx9c').getSheetByName('Customer List')
  const customerList_LodgeAndCharterSS = customerListSheet_LodgeAndCharterSS.getSheetValues(2, 1, customerListSheet_LodgeAndCharterSS.getLastRow() - 1, 2);
  var customerIndex, sheetName, sheetNameSplit;

  if (SpreadsheetApp.getActiveSheet().getSheetName() === 'Dashboard')
  {
    SpreadsheetApp.getActiveRangeList().getRanges().reverse().map(range => { 
      if (range.getColumn() === 3 && range.getLastColumn() === 3 && range.getRow() > 3)
      {
        range.getValues().map(customerName => {
          customerIndex = customerList.findIndex(customer => customer[1] === customerName[0]);
          sheetName = customerList[customerIndex][2]
          sheetNameSplit = sheetName.split(' - ')
          customerList.splice(customerIndex, 1);
          customerList_LodgeAndCharterSS.splice(customerList_LodgeAndCharterSS.findIndex(customer => customer[1] === customerName[0]), 1);
          spreadsheet.deleteSheet(spreadsheet.getSheetByName(sheetName))
          spreadsheet.deleteSheet(spreadsheet.getSheetByName(sheetNameSplit[0] + ' CHART - ' + sheetNameSplit[1]))
        })

        spreadsheet.deleteRows(range.getRow(), range.getNumRows())
      }
      else
        throw "You must select customer name(s)"
    })

    customerListSheet.getRange(3, 1, customerList.length, 3).setValues(customerList)
    customerListSheet_LodgeAndCharterSS.getRange(2, 1, customerList_LodgeAndCharterSS.length, 2).setValues(customerList_LodgeAndCharterSS)
  }
  else
  {
    spreadsheet.getSheetByName('Dashboard').activate()
    Browser.msgBox('You must be on the Dashboard to run this function.')
  }
}

/**
 * This function checks if the given string is blank or not.
 * 
 * @param {String} str : The given string
 * @return {Boolean} Whether the given string is blank or not.
 * @author Jarren Ralf
 */
function isNotBlank(str)
{
  return str !== ''
}

/**
 * This function process the imported data.
 * 
 * @param {Event Object} : The event object on an spreadsheet edit.
 * @author Jarren Ralf
 */
function processImportedData(e)
{
  if (e.changeType === 'INSERT_GRID')
  {
    var spreadsheet = e.source;
    var sheets = spreadsheet.getSheets();
    var info, numRows = 0, numCols = 1, maxRow = 2, maxCol = 3, isYearlyCustomerItemData = 4;

    for (var sheet = sheets.length - 1; sheet >= 0; sheet--) // Loop through all of the sheets in this spreadsheet and find the new one
    {
      if (sheets[sheet].getType() == SpreadsheetApp.SheetType.GRID) // Some sheets in this spreadsheet are OBJECT sheets because they contain full charts
      {
        info = [
          sheets[sheet].getLastRow(),
          sheets[sheet].getLastColumn(),
          sheets[sheet].getMaxRows(),
          sheets[sheet].getMaxColumns(),
          sheets[sheet].getRange(1, 5).getValue().toString().includes('Quantity Specif')
        ]
      
        // A new sheet is imported by File -> Import -> Insert new sheet(s) - The left disjunct is for a csv and the right disjunct is for an excel file
        if ((info[maxRow] - info[numRows] === 2 && info[maxCol] - info[numCols] === 2) || 
            (info[maxRow] === 1000 && info[maxCol] === 26 && info[numRows] !== 0 && info[numCols] !== 0) ||
            info[isYearlyCustomerItemData]) 
        {
          spreadsheet.toast('Processing imported data...', '', 60)
          const values = sheets[sheet].getSheetValues(1, 1, info[numRows], info[numCols]); 
          const sheetName = sheets[sheet].getSheetName()
          const sheetName_Split = sheetName.split(' ')
          const doesPreviousSheetExist = sheetName_Split[1]
          var fileName = sheetName_Split[0]

          if (sheets[sheet].getSheetName().substring(0, 7) !== "Copy Of") // Don't delete the sheets that are duplicates
            spreadsheet.deleteSheet(sheets[sheet]) // Delete the new sheet that was created

          if (info[isYearlyCustomerItemData])
            configureYearlyCustomerItemData(values, fileName, doesPreviousSheetExist, spreadsheet)

          spreadsheet.toast('', 'Import Complete.')
          break;
        }
      }
    }

    // Try and find the file created and delete it
    var file1 = DriveApp.getFilesByName(fileName + '.xlsx')
    var file2 = DriveApp.getFilesByName("Book1.xlsx")

    if (file1.hasNext())
      file1.next().setTrashed(true)

    if (file2.hasNext())
      file2.next().setTrashed(true)
  }
}

/**
 * This function protects all sheets expect for the search pages on the Lodge, Charter, & Guide data spreadsheet, for those, just the relevant cells in the header are protected.
 * 
 * @author Jarren Ralf
 */
function protectAllSheets()
{
  const users = ['triteswarehouse@gmail.com', 'scottnakashima10@gmail.com', 'scottnakashima@hotmail.com', 'pntparksville@gmail.com', 'derykdawg@gmail.com'];
  var sheetName, chartSheet = SpreadsheetApp.SheetType.OBJECT;

  SpreadsheetApp.getActive().getSheets().map(sheet => {
    if (sheet.getType() !== chartSheet)
    {
      sheetName = sheet.getSheetName();
      Logger.log(sheetName)

      if (sheetName !== 'Search for Item Quantity or Amount ($)')
      {
        if (sheetName !==  'Search for Invoice #s')
          sheet.protect().addEditor('jarrencralf@gmail.com').removeEditors(users);
        else
          sheet.protect().addEditor('jarrencralf@gmail.com').removeEditors(users).setUnprotectedRanges([sheet.getRange(1, 1, 2)]);
      }
      else
        sheet.protect().addEditor('jarrencralf@gmail.com').removeEditors(users).setUnprotectedRanges([sheet.getRange(1, 1, 3), sheet.getRange(2, 5, 2), sheet.getRange(2, 9, 2), sheet.getRange(3, 11)]);
      }
  })
}

/**
 * This function removes all of the protections on each sheet.
 * 
 * @author Jarren Ralf
 */
function removeProtectionOnAllSheets()
{
  var chartSheet = SpreadsheetApp.SheetType.OBJECT;

  SpreadsheetApp.getActive().getSheets().map(sheet => {
    if (sheet.getType() !== chartSheet)
      sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0].remove()
  })
}

/**
 * This function returns the user to the Dashboard sheet.
 * 
 * @author Jarren Ralf
 */
function returnToDashboard()
{
  SpreadsheetApp.getActive().getSheetByName('Dashboard').activate()
}

/**
 * This function spaces out the data and groups it by customer.
 * 
 * @param {String[][]} preData : The preformatted data.
 * @return {String[][], String[]} The reformatted data and a list of ranges to create a RangeList object
 * @author Jarren Ralf
 */
function reformatData_YearlyCustomerItemData(preData)
{
  var qty = 0, amount = 0, row = 0, uniqueCustomerList = [], ranges = [], formattedData = [];

  preData.map((customer, i, previousCustomers) => {
    if (uniqueCustomerList.includes(customer[0])) // Multiple Lines of Same Customer
    {
      qty += customer[4]
      amount += customer[5]
      formattedData.push(customer)
    }
    else if (uniqueCustomerList.length === 0) // First Customer
    {
      qty += customer[4]
      amount += customer[5]
      formattedData.push(customer)
      uniqueCustomerList.push(customer[0])
    }
    else // New Customer
    {
      formattedData.push([previousCustomers[i - 1][0], previousCustomers[ i -1][1], '', '', qty, amount], new Array(6).fill(''), customer)
      row = formattedData.length - 1;
      qty = customer[4];
      amount = customer[5];
      ranges.push('E' + row + ':F' + row)
      uniqueCustomerList.push(customer[0])
    }
  })

  const ii = preData.length - 1;

  // We need to add a row of totals for the final customer
  formattedData.push([preData[ii][0], preData[ii][1], '', '', qty, amount])
  row = formattedData.length + 1;
  ranges.push('E' + row + ':F' + row)

  return [formattedData, ranges]
}

/**
 * This function updates the sheet links on the dashboard
 * 
 * @author Jarren Ralf
 */
function setCustomerNamesOnDashboard()
{
  const spreadsheet = SpreadsheetApp.getActive()
  const dashboard = spreadsheet.getSheetByName('Dashboard')
  const customerListSheet = spreadsheet.getSheetByName('Customer List')
  const customerList = customerListSheet.getSheetValues(3, 1, customerListSheet.getLastRow() - 2, 2)
  const numRows = dashboard.getLastRow() - 3;
  const customerNames = dashboard.getSheetValues(4, 2, numRows, 1).map(acctNum => [customerList.find(customer => customer[0] == acctNum[0])[1]])
  dashboard.getRange(4, 3, numRows, 1).setValues(customerNames)
}

/**
 * This function updates the sheet links on the dashboard
 * 
 * @author Jarren Ralf
 */
function setSheetLinksOnDashboard()
{
  const sheets = SpreadsheetApp.getActive().getSheets();
  const dashboard = sheets.shift()
  const sheetNames = sheets.map(sheet => sheet.getSheetName().split(' - '))
  const numRows = dashboard.getLastRow() - 3

  const sheetLinks = dashboard.getSheetValues(4, 1, numRows, 1).map(custNum => {
    for (var s = 3; s < sheetNames.length; s++)
    {
      if (custNum[0] === sheetNames[s][1])
      {
        return [
          SpreadsheetApp.newRichTextValue().setText(custNum[0]).setLinkUrl('#gid=' + sheets[s    ].getSheetId()).build(),
          SpreadsheetApp.newRichTextValue().setText(custNum[0]).setLinkUrl('#gid=' + sheets[s + 1].getSheetId()).build() 
        ]
      }
    }
  })

  dashboard.getRange(4, 1, numRows, 2).setRichTextValues(sheetLinks)
}

/**
 * This function sets all of the tab colours so that the customer data tabs and charts are alternating colours for quick differentiation.
 * 
 * @author Jarren Ralf
 */
function setTabColours()
{
  var sheetNameSplit;

  SpreadsheetApp.getActive().getSheets().map(sheet => {
    sheetNameSplit = sheet.getSheetName().split(' - ');
    if (sheetNameSplit.length === 2)
      (sheetNameSplit[0].split(' ').pop() !== 'CHART') ? sheet.setTabColor('#38761d') : sheet.setTabColor('#f1c232');
  })
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

/**
 * This function deletes and rebuilds all of the charts in the spreadsheet in order to update the subtitle of the graph, which is the total Sales for a particular customer.
 * This function also contains the feature that if runtime is going to exceed 6 minutes, the limit for google apps script, then the script creates a trigger that will re-run
 * this function a few minutes later. This function creates the spreadsheets in a for-loop and if runtime will exceed 6 minutes, it stores the current value of the loop's 
 * incrementing variable in Google's CacheService, which stores string data that will expire after 6 minutes. On rerun, the function can call on the cache and resume within
 * the for-loop where the script was last stopped.
 * 
 * @author Jarren Ralf
 */
function updateAllCharts()
{
  try
  {
    const startTime = new Date(); // The start time of this function
    const MAX_RUNNING_TIME = 330000; // Five minutes thirty seconds
    var REASONABLE_TIME_TO_WAIT = 30000; // Thirty seconds
    const spreadsheet = SpreadsheetApp.getActive()
    const sheets = spreadsheet.getSheets();
    const dashboard = sheets.shift()
    const numRows = dashboard.getLastRow() - 3
    const totalYearlySalesPerCustomer = dashboard.getSheetValues(4, 3, numRows, 2).map(total => [total[0], twoDecimals(total[1])])
    const sheetNames = sheets.map(sheet => sheet.getSheetName().split(' - '));
    const numYears = new Date().getFullYear() - 2011;
    const numCustomerSheets = sheetNames.length - numYears - 1
    var cache = CacheService.getDocumentCache(), customerIndex = 0, chart, currentTime = 0;
    var currentSheet = Number(cache.get('current_sheet'));

    if (currentSheet === 0) // If the cache was null, set the initial sheet index to 3
      currentSheet = 3;

    // Create the spreadsheets, notice that the index varibale needs to be converted to a number since the Cache stores data as string values
    for (var sheet = currentSheet; sheet < numCustomerSheets; sheet = sheet + 2)
    {
      currentTime = new Date().getTime();
      
      if (currentTime - startTime >= MAX_RUNNING_TIME) // If the function has been running for more than 5 minutes, then set the trigger to run this function again in a few minutes
      {
        cache.put('current_sheet', sheet.toString()); // Store the indexing variable

        var triggerDate = new Date(currentTime + REASONABLE_TIME_TO_WAIT); // Set a trigger for a point in the future
        Logger.log('Next Trigger will run at:')
        Logger.log(triggerDate)

        ScriptApp.newTrigger("updateAllCharts").timeBased().at(triggerDate).create();
        break;
      }
      else
      {
        spreadsheet.deleteSheet(sheets[sheet + 1]) // Delete the chart

        chart = sheets[sheet].newChart()
          .asColumnChart()
          .addRange(sheets[sheet].getRange(3, 5, numYears, 2))
          .setNumHeaders(0)
          .setXAxisTitle('Year')
          .setYAxisTitle('Sales Total')
          .setTransposeRowsAndColumns(false)
          .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
          .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
          .setOption('title', totalYearlySalesPerCustomer[customerIndex][0])
          .setOption('subtitle', 'Total: $' + new Intl.NumberFormat().format(totalYearlySalesPerCustomer[customerIndex][1]))
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

        sheets[sheet].insertChart(chart);
        spreadsheet.moveChartToObjectSheet(chart).setName(sheetNames[sheet][0] + ' CHART - ' + sheetNames[sheet][1]).getSheetId()
        customerIndex++;
      }
    }

    if (sheet === numCustomerSheets) // The total number of spreadsheets have been created
      setSheetLinksOnDashboard()

    const salesDataSheet = spreadsheet.getSheetByName('Sales Data');
    const spreadsheetName = spreadsheet.getName();
    ss.deleteSheet(ss.getSheetByName('ANNUAL ' + spreadsheetName + ' CHART')) // Delete previous sales chart

    const annualSalesChart = salesDataSheet.newChart()
      .asColumnChart()
      .addRange(salesDataSheet.getRange(4, 1, numYears, 2))
      .setNumHeaders(0)
      .setXAxisTitle('Year')
      .setYAxisTitle('Sales Total')
      .setTransposeRowsAndColumns(false)
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('title', 'ANNUAL ' + spreadsheetName + ' DATA')
      .setOption('subtitle', 'Total: ' + salesDataSheet.getRange(2, 2).getDisplayValue())
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

    salesDataSheet.insertChart(annualSalesChart);
    spreadsheet.moveChartToObjectSheet(annualSalesChart).activate().setName('ANNUAL ' + spreadsheetName + ' CHART')

    const ss = SpreadsheetApp.openById('1xKw4GAtNbAsTEodCDmCMbPCbXUlK9OHv0rt5gYzqx9c') // The Lodge, Charter, & Guide Data spreadsheet
    const annualSalesDataSheet = ss.getSheetByName('Annual Sales Data');
    ss.deleteSheet(ss.getSheetByName('ANNUAL SALES CHART')) // Delete previous sales chart

    const annualSalesChart_BOTH = annualSalesDataSheet.newChart()
      .asColumnChart()
      .addRange(annualSalesDataSheet.getRange(4, 1, numYears, 2))
      .setNumHeaders(0)
      .setXAxisTitle('Year')
      .setYAxisTitle('Sales Total')
      .setTransposeRowsAndColumns(false)
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('title', 'Annual Sales Data')
      .setOption('subtitle', 'Total: ' + annualSalesDataSheet.getRange(2, 2).getDisplayValue())
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

    annualSalesDataSheet.insertChart(annualSalesChart_BOTH);
    spreadsheet.moveChartToObjectSheet(annualSalesChart_BOTH).activate().setName('ANNUAL SALES CHART')
  }
  catch (err)
  {
    var error = err['stack'];
    Logger.log(error)

    if (sheet !== numCustomerSheets)// If there are still more spreadsheets to create
    {
      var triggerDate = new Date(currentTime + REASONABLE_TIME_TO_WAIT); 
      
      Logger.log('Next Trigger will run at:')
      Logger.log(triggerDate)

      ScriptApp.newTrigger("updateAllCharts").timeBased().at(triggerDate).create(); // Create a new trigger and try running the function again
      cache.put('current_sheet', sheet.toString()); // Store the current position of the for-loop iterate
    }
  }
}

/**
 * This function looks through all of the sheets and updates the sales date for all years since 2012, for all customers. The function
 * finishes by updating the Dashboard with the sales data.
 * 
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @author Jarren Ralf
 */
function updateAllCustomersSalesData(spreadsheet)
{
  if (arguments.length === 0)
  {
    spreadsheet = SpreadsheetApp.getActive()
  }

  const today = new Date();
  const currentYear = today.getFullYear();
  const currentDate = today.getDate() + ' ' + ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][today.getMonth()] + ' ' + currentYear;
  const numYears = currentYear - 2012 + 1
  const sheets = spreadsheet.getSheets();
  const dashboard = sheets.shift()
  const sheetNames = sheets.map(sheet => sheet.getSheetName().split(' - '));
  const numCustomerSheets = sheetNames.length - numYears - 1
  const range = dashboard.getRange(4, 5, dashboard.getLastRow() - 3, dashboard.getLastColumn() - 4)
  const salesTotals = range.getValues();
  const hAligns = ['left', 'left', 'right', 'right'], numFormats = ['@', '@', '@', '$#,##0.00']
  const chartDataFormat = new Array(numYears).fill().map(() => ['@', '$#,##0.00']);
  const chartDataH_Alignment = new Array(numYears).fill().map(() => ['center', 'right']);
  var sheet, data, numItems = 0, chartData = [], index = 0, allYearsData, salesData, hAlignments = [], numberFormats = [], 
    yearRange = [], yearRange_RowNum = 3, totalRange = [], totalRange_RowNum = 1;

  const years = new Array(numYears).fill('').map((_, y) => (currentYear - y).toString()).map(year_y => {
    chartData.push([year_y, ''])
    sheet = spreadsheet.getSheetByName(year_y)
    return sheet.getSheetValues(2, 1, sheet.getLastRow() - 1, 6)
  })

  chartData.reverse() 

  for (var s = 3; s < numCustomerSheets; s = s + 2)
  {
    spreadsheet.toast((index + 1) + ': ' + sheetNames[s][0] + ' - ' + sheetNames[s][1], 'Updating...', 60)
    
    allYearsData = years.map((fullYearData, y) => {
      data = fullYearData.filter(custNum => custNum[0].trim() === sheetNames[s][1])
      numItems = data.length;

      if (numItems !== 0)
      {
        chartData[numYears - y - 1][1] = data[numItems - 1][5];
        salesTotals[index][y] = data[numItems - 1][5]; 
        ((currentYear - y) == currentYear) ? 
          data.unshift(['', '', '', '', '01 Jan ' + currentYear, currentDate]) : 
          data.unshift(['', '', '', '', '01 Jan ' + (currentYear - y), '31 Dec ' + (currentYear - y)])
        data.push(['', '', '', '', '', '']);
        totalRange_RowNum += (totalRange_RowNum == 0) ? numItems + 1 : numItems + 2;
        totalRange.push('C' + totalRange_RowNum + ':D' + totalRange_RowNum)
        yearRange.push('C' + yearRange_RowNum + ':D' + yearRange_RowNum)
        yearRange_RowNum += numItems + 2;
      }
      else
      {
        chartData[numYears - y - 1][1] = ''
        salesTotals[index][y] = ''; 
      }

      return data.map(col => [col[2], col[3], col[4], col[5]])
    })

    index++

    salesData = [].concat.apply([], allYearsData);
    salesData.pop()

    hAlignments = new Array(salesData.length).fill().map(() => hAligns)
    numberFormats = new Array(salesData.length).fill().map(() => numFormats)

    sheets[s].getRange(3, 1, sheets[s].getMaxRows() - 2, 6).clearContent().setBackground('white').setBorder(false, false, false, false, false, false)
      .offset(0, 0, salesData.length, 4).setFontWeight('normal').setVerticalAlignment('middle').setHorizontalAlignments(hAlignments).setNumberFormats(numberFormats).setValues(salesData)
      .offset(0, 4, numYears, 2).setNumberFormats(chartDataFormat).setHorizontalAlignments(chartDataH_Alignment).setFontWeight('normal').setValues(chartData)
      .offset(-2, -1, 1, 1).setFormula([['=SUM(F3:F' + (numYears + 2) + ')']])

    sheets[s].getRangeList(yearRange).setFontWeight('bold').setNumberFormat('@') // The year
    sheets[s].getRangeList(totalRange).setBorder(true, false, true, false, false, false).setBackground('#c0c0c0').setFontWeight('bold') // The total quantity and amount

    yearRange.length = 0;
    totalRange.length = 0;
    hAlignments.length = 0;
    numberFormats.length = 0;
    yearRange_RowNum = 3;
    totalRange_RowNum = 1;
  }

  const yearlySales = range.setNumberFormat('$#,##0.00').setValues(salesTotals).activate().offset(-1, 0, 1, numYears).getDisplayValues()[0].reverse();
  const annualSalesData = [];

  if (spreadsheet.getName().split(' ', 1)[0] !== 'CHARTER')
  {
    var charterGuideSalesYearlyData = SpreadsheetApp.openById('1kKS6yazOEtCsH-QCLClUI_6NU47wHfRb8CIs-UTZa1U').getSheetByName('Sales Data').getDataRange().getDisplayValues();
    charterGuideSalesYearlyData.shift()
    charterGuideSalesYearlyData.shift()
    charterGuideSalesYearlyData.shift()

    var annualChartData = yearlySales.map((total, y) => {
      annualSalesData.push([(2012 + y).toString(), '=SUM(C' + (y + 4) + ':D' + (y + 4) + ')', 
        total, (charterGuideSalesYearlyData[y] != null) ? charterGuideSalesYearlyData[y][1] : '$0.00'])
      return [(2012 + y).toString(), total]
    });
  }
  else
  {
    var lodgeSalesYearlyData = SpreadsheetApp.openById('1o8BB1RWkxK1uo81tBjuxGc3VWArvCdhaBctQDssPDJ0').getSheetByName('Sales Data').getDataRange().getDisplayValues();
    lodgeSalesYearlyData.shift()
    lodgeSalesYearlyData.shift()
    lodgeSalesYearlyData.shift()

    var annualChartData = yearlySales.map((total, y) => {
      annualSalesData.push([(2012 + y).toString(), '=SUM(C' + (y + 4) + ':D' + (y + 4) + ')', 
        (lodgeSalesYearlyData[y] != null) ? lodgeSalesYearlyData[y][1] : '$0.00', total])
      return [(2012 + y).toString(), total]
    })
  }

  SpreadsheetApp.openById('1xKw4GAtNbAsTEodCDmCMbPCbXUlK9OHv0rt5gYzqx9c').getSheetByName('Annual Sales Data').getRange(4, 1, numYears, 4).setValues(annualSalesData)
  spreadsheet.getSheetByName('Sales Data').getRange(4, 1, numYears, 2).setNumberFormats(chartDataFormat).setValues(annualChartData)

  var triggerDate = new Date(new Date().getTime() + 30000); // Set a trigger for a point in the future
  Logger.log('All of the charts will begin updating at:')
  Logger.log(triggerDate)      
  ScriptApp.newTrigger("updateAllCharts").timeBased().at(triggerDate).create();
  spreadsheet.toast('', 'Full Data Update: COMPLETE', 60)
}