/**
 * This function handles the onChange trigger event. It checks 
 * 
 * @param {Event} e : The event object 
 */
function onChange(e)
{
  if (e.changeType === 'INSERT_GRID') // A new sheet has been created
    processImportedData(e)
  else if (e.changeType === 'REMOVE_ROW') // A row has been deleted
    reformatPackingSlip(e)
}

/**
 * This function handles the onEdit trigger event. It checks ...
 * 
 * @param {Event} e : The event object 
 */
function onEdit(e)
{
  var range = e.range; 
  var row = range.rowStart;
  var col = range.columnStart;

  if (row == range.rowEnd && col == range.columnEnd) // Single cell
  {
    var spreadsheet = e.source;
    var sheet = spreadsheet.getActiveSheet();
    var sheetName = sheet.getSheetName();
    var checkboxRange = spreadsheet.getRangeByName('Checkboxes');

    /* 
    * If a single cell in the range of the checkboxes is being edited, then observe if any of the checkboxes are checked. 
    * When these checkboxes are checked, their values correspond to the appropriate taxation value, eg. BC: 0.12, AB: 0.05, USA: 0, etc. and 
    * they are false otherwise. Therefore, any value greater than or equal to 0 represents a checked box. Leave the box that was immediately 
    * clicked checked, which is identified by the row that was edited minus 3 (because the top of the range starts at row 3).
    */
    if (sheetName === 'Calculator'  && col == 4 && row > 2 && row < 9)
      checkboxRange.setValues(checkboxRange.getValues().map((check, index) => (check[0] >= 0 && index != row - 3) ? [false] : check))
    else if (sheetName === 'Status Page' && col == 2 && row > 1)
    {
      range.uncheck()
      const ui = SpreadsheetApp.getUi();
      const orderNum = range.offset(0, -1).getValue()
      const fullData = spreadsheet.getSheetByName('All_Active_Orders').getDataRange().getValues()
      const header = fullData.shift()
      const shopifyData = fullData.filter(val => val[0] === orderNum);
      shopifyData.unshift(header)

      if (shopifyData.length === 1)
        ui.alert('Order Not Found', 'Use File -> Import to generate a Invoice for ' + orderNum + '.', ui.ButtonSet.OK)
      else
      {
        updatePackingSlip(shopifyData, shopifyData.length, shopifyData[0].length, spreadsheet)
        spreadsheet.getSheetByName('Invoice').getRange('I5').activate();
      }
    }
    else if (sheetName === 'Invoice')
    {
      const numRowsPerPage = 49;

      if (row == 4 && col == 5) // Hidden checkbox that adds (or removes) 10% to the freight cost
      {
        const rng = spreadsheet.getRangeByName('ShippingAmount');
        var shippingAmount = rng.getValue()
        shippingAmount = (range.isChecked()) ? twoDecimals(shippingAmount*1.1) : twoDecimals(shippingAmount/1.1)
        rng.setValue(shippingAmount)
      }
      else if (row == 5 && col == 5) // Hidden checkbox that removes taxes from the order
      {
        const orderNum = range.offset(-4, 4).getValue()
        const fullData = spreadsheet.getSheetByName('All_Active_Orders').getDataRange().getValues()
        fullData.shift()
        const shopifyData = fullData.find(val => val[0] === orderNum);
        const checks = checkboxRange.getValues()

        // Check the shipping country and province, then set the taxes accordingly by checking the appropriate box
        if (shopifyData[0][41] == null) // Blank means the item is a pick up in BC, therefore charge 12%
        {
          checks[0][0] = 0.12;
          spreadsheet.getRangeByName('ShippingAmount').setValue(0);
        }
        else
        {
          if (shopifyData[0][42] !== 'CA')
            checks[5][0] = 0;
          else
          {
            if (shopifyData[0][41] === 'BC') 
              checks[0][0] = 0.12;
            else if (shopifyData[0][41] === 'AB' || shopifyData[0][41] === 'NT' || shopifyData[0][41] === 'NU' || 
                    shopifyData[0][41] === 'YT' || shopifyData[0][41] === 'QC' || shopifyData[0][41] === 'MB')
              checks[1][0] = 0.05;
            else if (shopifyData[0][41] === 'NS' || shopifyData[0][41] === 'NB' || shopifyData[0][41] === 'NL' || shopifyData[0][41] === 'PE')
              checks[2][0] = 0.15;
            else if (shopifyData[0][41] === 'ON')
              checks[3][0] = 0.13;
            else if (shopifyData[0][41] === 'SK')
              checks[4][0] = 0.11;
          }
        }

        if (range.isChecked())
          checkboxRange.uncheck()
        else 
        {
          checkboxRange.setValues(checks)
          sheet.getRange(5, 9).setFormula('=Items_Tax+Shipping_Tax')
        }
      }
      else if (row == 14) 
      {
        if (col == 2) // Dropdown box that selects the shipping method - Checking for Lettermail or Pick Up
        {
          switch (e.value)
          {
            case 'Post Lettermail':
              spreadsheet.getRangeByName('Checkbox_Lettermail').check()
              spreadsheet.getRangeByName('Checkbox_PickUp').uncheck()
              break;
            case 'Pick Up':
              spreadsheet.getRangeByName('ShippingAmount').setValue(0);
              spreadsheet.getRangeByName('Checkbox_PickUp').check()
              spreadsheet.getRangeByName('Checkbox_Lettermail').uncheck()
              //changeShippingAddress()
              break;
            case 'Post Tracked Packet':
            case 'Post Expedited Parcel':
            case 'Post Xpress Post':
            case 'Purolator Ground':
            case 'Purolator Express':
            case 'UPS Standard':
            case 'UPS Express':
              const trackingNumber = sheet.getRange(14, 4).getValue();

              if (isNotBlank(trackingNumber))
              {
                const linkUrl = getTrackingNumberURL(sheet.getRange(14, 4).getValue(), e.value)
                const hyperlink = SpreadsheetApp.newTextStyle().setFontSize(11).setUnderline(true).setForegroundColor('#1155cc').build();
                const hyperlinkRichText = SpreadsheetApp.newRichTextValue().setText(trackingNumber).setTextStyle(hyperlink).setLinkUrl(linkUrl).build();
                sheet.getRange(14, 4).setRichTextValue(hyperlinkRichText)
              }
              break;
            default:
              spreadsheet.getRangeByName('Checkbox_PickUp').uncheck()
              spreadsheet.getRangeByName('Checkbox_Lettermail').uncheck()
          }
        }
        else if (col == 4) // The cell that has the tracking number
        {
          const trackingNumber = range.getValue();
          const linkUrl = getTrackingNumberURL(trackingNumber, sheet.getRange(14, 2).getValue())
          const hyperlink = SpreadsheetApp.newTextStyle().setFontSize(11).setUnderline(true).setForegroundColor('#1155cc').build();
          const hyperlinkRichText = SpreadsheetApp.newRichTextValue().setText(trackingNumber).setTextStyle(hyperlink).setLinkUrl(linkUrl).build();
          range.setBorder(true, false, true, false, false, false).setRichTextValue(hyperlinkRichText)
        }
      }
      else if ((row > 16 && row < 49) || (row % numRowsPerPage >= 10 && row % numRowsPerPage <= 48))
      {
        if (col == 1) // Quantity Changed
        {
          const rng = sheet.getRange(row, 1, 1, 9);
          const vals = rng.getValues();
          vals[0][0] = e.value + ' x'
          vals[0][8] = vals[0][7]*e.value;
          rng.setValues(vals);
        }
        else if (col == 8) // Unit price changed
        {
          const rng = sheet.getRange(row, 1, 1, 9);
          const vals = rng.getValues();
          vals[0][8] = e.value*vals[0][0].split(' ', 1)[0];
          rng.setValues(vals);
        }
      }
    }
  }
}

/**
 * This function handles the onOpen trigger event. It checks ...
 * 
 * @param {Event} e : The event object 
 */
function onOpen()
{
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Invoice Controls')
    .addItem('COMPLETE', 'packingSlip_Complete')
    .addItem('Hold For Pick Up', 'packingSlip_HFPU')
    .addItem('Back Order', 'packingSlip_BackOrder')
    .addSubMenu(ui.createMenu('Show Items')
    .addItem('ALL Items', 'showItems_all')
    .addItem('Pending ONLY', 'showItems_pending')
    .addItem('Fulfilled ONLY', 'showItems_fulfilled')
    .addItem('Unfulfilled ONLY', 'showItems_unfulfilled')
    .addItem('Pending & Fulfilled', 'showItems_pending_AND_fulfilled'))
    //.addItem('Display Shipping Calculator', 'displayShippingCalculator')
    .addItem('Email Invoice To Customer', 'emailPackingSlip')
    .addSeparator()
    .addItem('Apply Formatting', 'applyFormatting')
    //.addItem('Display Shipping Calculator', 'displayShippingCalculator')
    .addSeparator()
    .addItem('Clear Export Page', 'clearExportPage')
    .addSeparator()
    .addItem('Download Packing Slip and Invoice', 'downloadButton')
    .addToUi();

  resetArrayFormulaPackingSlip();
  applyFormatting()
}

/**
 * This function loops through all of the sheets in the spreadsheet and formats them. If an argument is passed to the function in the
 * form of an array of Sheet objects, then the fucntion will apply formatting to only those specific sheets.
 * 
 * @param {Sheet[]} sheets : The set of sheets in the spreadsheet to be formatted.
 * @author Jarren Ralf
 */
function applyFormatting(sheets)
{
  const spreadsheet = SpreadsheetApp.getActive();

  if (arguments.length === 0) 
    sheets = spreadsheet.getSheets();

  var sheetName = '', range, lastRow, lastCol;

  for (var s = 0; s < sheets.length; s++)
  {
    sheetName = sheets[s].getSheetName();

    if (sheetName === 'Status Page')
    {
      range = sheets[s].setColumnWidth(1, 75).setColumnWidth(2, 30).setColumnWidth(3, 126).setColumnWidth(4, 331)
        .setColumnWidth(5, 117).setColumnWidth( 6, 122).setColumnWidth( 7, 118).setColumnWidth( 8, 248)
        .setColumnWidth(9, 108).setColumnWidth(10, 120).setColumnWidth(11, 118).setColumnWidth(12,  93).getDataRange();
      lastRow = range.getLastRow()
      lastCol = range.getLastColumn()
      sheets[s].setFrozenRows(2)

      const conditionalFormattingColours = sheets[s].getConditionalFormatRules().map(rule => rule.getBooleanCondition().getBackgroundObject().asRgbColor().asHexString());
      const colours = range.getBackgrounds().map((row, i) => row.map(colour => (conditionalFormattingColours.includes(colour) && i > 1) ? '#ffffff' : colour))

      range.setFontColor('black').setFontFamily('Calibri').setVerticalAlignment('middle').setBackgrounds(colours).setFontStyle('normal').setFontWeight('normal')
        .setNumberFormats(new Array(lastRow).fill(['@', '0.###############', ...new Array(lastCol - 3).fill('@'), 'dd MMM yyyy']))
        .setFontSizes([new Array(lastCol).fill(18), ...new Array(lastRow - 1).fill(new Array(lastCol).fill(12))])
        .setHorizontalAlignments([['center', ...new Array(lastCol - 1).fill('left')], new Array(lastCol).fill('left'), 
          ...new Array(lastRow - 2).fill(['left', 'middle', ...new Array(lastCol - 3).fill('left'), 'right'])])
    }
    else if (sheetName === 'Invoice')
    {
      const col = 9; // Number of columns on the packing slip
      const numRowsPerPage = 49;
      const numItemsOnPageOne = 32;
      const numItemsPerPage = 39; // Starting with page 2
      const pageNumber = sheets[s].getRange(numRowsPerPage + 1, col).getValue()
      const numPages = (isBlank(pageNumber)) ? 1 : pageNumber.split(' of ')[1];
      const numRows = numRowsPerPage*numPages + 1;
      const values = sheets[s].getSheetValues(1, 6, 8, 4);
      const shippingMethod = sheets[s].getSheetValues(14, 2, 1, 1)[0][0]
      const ordNumber = values[0][3]
      const shippingAmount = values[3][3]
      const customerName = values[7][0]
      var N;

      var subtotalAmount = '=SUM(Item_Totals_Page_1';

      for (var n = 2; n <= numPages; n++)
        subtotalAmount += ',' + 'Item_Totals_Page_' + n

      subtotalAmount += ')';

      const pntLogo = SpreadsheetApp.newCellImage().toBuilder().setSourceUrl('http://cdn.shopify.com/s/files/1/0018/7079/0771/files/logoh_180x@2x.png?v=1613694206').build();

      const pntAddress = SpreadsheetApp.newRichTextValue().setText('3731 Moncton Street, Richmond, BC, V7E 3A5\nPhone: (604) 274-7238 Toll Free: (800) 895-4327\nwww.pacificnetandtwine.com')
        .setLinkUrl(91, 117, 'https://www.pacificnetandtwine.com/').build()

      const boldTextStyle = SpreadsheetApp.newTextStyle().setBold(true).setFontSize(12).build();
      const shipDate = SpreadsheetApp.newRichTextValue().setText('Ship Date: ' + Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), "dd MMMM yyyy"))
        .setTextStyle(0, 10, boldTextStyle).build();

      const emailHyperLink = SpreadsheetApp.newRichTextValue().setText('If you have any questions, please send an email to: websales@pacificnetandtwine.com')
        .setLinkUrl(52, 83, 'mailto:websales@pacificnetandtwine.com?subject=RE: [Pacific Net %26 Twine Ltd] ' + ordNumber + ' placed by ' + customerName).build()

      sheets[s].setColumnWidth(1, 40).setColumnWidth(2, 133).setColumnWidth(3, 162).setColumnWidth(4, 156)
          .setColumnWidth(5, 40).setColumnWidth(6, 95).setColumnWidth(7, 50).setColumnWidths(8, 2, 75)
          .setRowHeight(1, 20).setRowHeight(2, 40).setRowHeights(3, 4, 20).setRowHeight(7, 10).setRowHeights(8, 5, 20).setRowHeight(13, 10)
          .setRowHeight(14, 20).setRowHeight(15, 10).setRowHeights(16, 33, 20).setRowHeight(49, 10).setRowHeight(50, 20)
        .getRange(1, 1) // PNT Logo
          .setValue(pntLogo)
        .offset(0, 7, 6, 2) // The invoice header values on page one
          .setFontColor('black').setFontFamily('Arial').setFontStyle('normal').setBackground('white')
          .setFontSizes([[10, 10], [10, 9], [10, 10], [10, 10], [10, 10], [10, 10]])
          .setVerticalAlignments([['middle', 'middle'], ['top', 'top'], ['middle', 'middle'], ['middle', 'middle'], ['middle', 'middle'], ['middle', 'middle']])
          .setNumberFormats([['@', '@'], ['@', 'dd MMM yyyy'], ['@', '$#,##0.00'], ['@', '$#,##0.00'], ['@', '$#,##0.00'], ['@', '$#,##0.00']])
          .setHorizontalAlignments([['right', 'center'], ['right', 'center'], ['right', 'right'], ['right', 'right'], ['right', 'right'], ['right', 'right']])
          .setFontWeights(new Array(6).fill(['bold', 'normal']))
          .setValues([['Web Order Number', '=Last_Import!A2'], ['Ordered Date', '=INDEX(SPLIT(Last_Import!P2, " "), 1, 1)'], 
            ['Subtotal Amount:', subtotalAmount], ['Shipping Amount:', shippingAmount], 
            ['Taxes:', '=Items_Tax+Shipping_Tax'], ['Order Total:', '=SUM(Order_Subtotals)']])
        .offset(3, -7) // PNT Address
          .setRichTextValue(pntAddress)
        .offset(4, 0, 5) // The value "SHIP" in the header of the packing slip
          .setBackground('#d9d9d9').setBorder(true, true, true, true, false, false).setFontColor('black').setFontFamily('Arial')
          .setFontLine('none').setFontSize(14).setFontStyle('normal').setFontWeight('bold').setHorizontalAlignment('center').setNumberFormat('@')
          .setVerticalAlignment('middle').setVerticalText(true).setValue('SHIP')
        .offset(0, 1, 5, 2) // The "SHIP" values
          .mergeAcross().setBackground('white').setBorder(true, true, true, true, false, false).setFontColor('black').setFontFamily('Arial')
          .setFontLine('none').setFontSize(10).setFontStyle('normal').setFontWeight('normal').setHorizontalAlignment('left').setNumberFormat('@')
          .setVerticalAlignment('middle').setVerticalText(false).setFormulas([['Last_Import!AI2', null], ['Last_Import!AM2', null], 
            ['Last_Import!AK2', null], ['Last_Import!AN2&", "&Last_Import!AP2&", "&Last_Import!AO2&", "&Last_Import!AQ2', null], ['Last_Import!AR2', null]])
        .offset(0, 3, 5, 1) // The value "BILL" in the header of the packing slip
          .setBackground('#d9d9d9').setBorder(true, true, true, true, false, false).setFontColor('black').setFontFamily('Arial')
          .setFontLine('none').setFontSize(14).setFontStyle('normal').setFontWeight('bold').setHorizontalAlignment('center').setNumberFormat('@')
          .setVerticalAlignment('middle').setVerticalText(true).setValue('BILL')
        .offset(0, 1, 5, 4) // The "BILL" values
          .mergeAcross().setBackground('white').setBorder(true, true, true, true, false, false).setFontColor('black').setFontFamily('Arial')
          .setFontLine('none').setFontSize(10).setFontStyle('normal').setFontWeight('normal').setHorizontalAlignment('left').setNumberFormat('@')
          .setVerticalAlignment('middle').setVerticalText(false).setFormulas([['Last_Import!Y2', null, null, null], ['Last_Import!AC2', null, null, null], ['Last_Import!AA2', null, null, null], 
            ['Last_Import!AD2&", "&Last_Import!AF2&", "&Last_Import!AE2&", "&Last_Import!AG2', null, null, null], ['Last_Import!AH2', null, null, null]])
        .offset(6, -5, 1, col) // The shipping values
          .setBackground('white').setBorder(true, true, true, true, false, false).setFontColor('black').setFontFamily('Arial')
          .setFontLine('none').setFontStyle('normal').setNumberFormat('@').setVerticalAlignment('middle').setVerticalText(false)
          .setFontSizes([[12, ...new Array(col - 1).fill(10)]])
          .setFontWeights([['bold', ...new Array(col - 1).fill('normal')]])
          .setHorizontalAlignments([[...new Array(col - 3).fill('left'), 'right', 'left', 'left']])
          .setValues([['Via', shippingMethod, ...new Array(col - 2).fill('')]])
        .offset(0, 1, 1, 2).merge() // Shipping method cells
        .offset(0, 5, 1, 3).merge() // Ship Date cells
          .setRichTextValue(shipDate)
        .offset(2, -4, numItemsOnPageOne + 1, 5) // The item column
          .mergeAcross()
        .offset(0, -2, numItemsOnPageOne + 1, col) // All the item information
            .setFontColor('black').setFontFamily('Arial').setFontStyle('normal').setVerticalAlignment('middle').setBackground('white')
            .setFontWeights([new Array(col).fill('bold'), ...new Array(numItemsOnPageOne).fill(new Array(col).fill('normal'))])
            .setFontSizes([new Array(col).fill(12), ...new Array(numItemsOnPageOne).fill(new Array(col).fill(9))])
            .setNumberFormats([new Array(col).fill('@'), ...new Array(numItemsOnPageOne).fill([...new Array(col - 2).fill('@'), '$#,##0.00', '$#,##0.00'])])
            .setHorizontalAlignments(new Array(numItemsOnPageOne + 1).fill(['center', 'center', 'left', 'left', 'left', 'left', 'left', 'right', 'right']))
            .setBorder(true, true, true, true, true, false)
        .offset(numItemsOnPageOne + 2, 0, 1, 5) // The hyperlinked email at the bottom of the page
          .merge().setRichTextValue(emailHyperLink)
        
      for (var n = 0; n < numPages - 1; n++)
      {
        N = numRowsPerPage*n;

        sheets[s].setRowHeight(51 + N, 20).setRowHeight(52 + N, 40).setRowHeights(53 + N, 4, 20)
            .setRowHeight(57 + N, 10).setRowHeights(58 + N, 40, 20).setRowHeight(98 + N, 10).setRowHeight(99 + N, 10)
          .getRange(50 + N, col - 1, 7, 2) // The invoice header values on each page
            .setFontColor('black').setFontFamily('Arial').setFontStyle('normal').setBackground('white')
            .setFontSizes([[10, 10], [10, 10], [10, 9], [10, 10], [10, 10], [10, 10], [10, 10]])
            .setVerticalAlignments([['middle', 'middle'], ['middle', 'middle'], ['top', 'top'], ['middle', 'middle'], ['middle', 'middle'], ['middle', 'middle'], ['middle', 'middle']])
            .setNumberFormats([['@', '@'], ['@', '@'], ['@', 'dd MMM yyyy'], ['@', '$#,##0.00'], ['@', '$#,##0.00'], ['@', '$#,##0.00'], ['@', '$#,##0.00']])
            .setHorizontalAlignments([['left', 'right'], ['right', 'center'], ['right', 'center'], ['right', 'right'], ['right', 'right'], ['right', 'right'], ['right', 'right']])
            .setFontWeights(new Array(7).fill(['bold', 'normal']))
            .setValues([['', 'Page ' + (n + 1) + ' of ' + numPages], ['Web Order Number', '=I1'], ['Ordered Date', '=I2'], 
              ['Subtotal Amount:', '=I3'], ['Shipping Amount:', '=I4'], ['Taxes:', '=I5'], ['Order Total:', '=I6']])
          .offset(1, -7, 1, 1) // PNT Logo on each page
            .setValue(pntLogo)
          .offset(3, 0, 1, 1) // PNT Address on each page
            .setRichTextValue(pntAddress)
          .offset(4, 2, numItemsPerPage + 1, 5) // Item column on each page
            .mergeAcross()
          .offset(0, -2, numItemsPerPage + 1, col) // Item information on each page
            .setFontColor('black').setFontFamily('Arial').setFontStyle('normal').setVerticalAlignment('middle').setBackground('white')
            .setFontWeights([new Array(col).fill('bold'), ...new Array(numItemsPerPage).fill(new Array(col).fill('normal'))])
            .setFontSizes([new Array(col).fill(12), ...new Array(numItemsPerPage).fill(new Array(col).fill(9))])
            .setNumberFormats([new Array(col).fill('@'), ...new Array(numItemsPerPage).fill([...new Array(col - 2).fill('@'), '$#,##0.00', '$#,##0.00'])])
            .setHorizontalAlignments(new Array(numItemsPerPage + 1).fill(['center', 'center', 'left', 'left', 'left', 'left', 'left', 'right', 'right']))
            .setBorder(true, true, true, true, true, false)
          .offset(numItemsPerPage + 2, 0, 1, 5) // Email hyperlink on each page
            .merge().setRichTextValue(emailHyperLink)
      }

      if (numPages > 1)
        sheets[s].getRange(numRows, col).setFontColor('black').setFontFamily('Arial').setFontStyle('normal').setBackground('white')
              .setFontSize(10).setVerticalAlignment('middle').setNumberFormat('@').setHorizontalAlignment('right').setFontWeight('normal')
              .setValue('Page ' + numPages + ' of ' + numPages)
      else
        sheets[s].getRange(numRows, col).setValue('')
    }
    else if (sheetName === 'Calculator')
    {
      range = sheets[s].setColumnWidth(1, 15).setColumnWidths(2, 2, 100).setColumnWidth(4, 35).setColumnWidth(5, 170).setColumnWidth(6, 15)
        .setRowHeightsForced(1, 1, 15).setRowHeights(2, 9, 25).setRowHeight(11, 15).getRange(1, 1, sheets[s].getMaxRows(), sheets[s].getMaxColumns());
      lastRow = range.getLastRow()
      lastCol = range.getLastColumn()

      range.setBackground('#a4c2f4').setFontFamily('Arial').setVerticalAlignment('middle')
        .setFontColor([['black', 'black', 'black', '#a4c2f4', 'black', 'black'], ['black', 'black', 'black', 'black', 'black', 'black'], 
          ['black', 'black', 'black', 'black', 'black', 'black'], ['black', 'black', '#a4c2f4', 'black', 'black', 'black'], 
          new Array(lastRow - 4).fill(new Array(lastCol).fill('black'))])
        .setFontSizes([...new Array(lastRow - 2).fill([12, 12, 12, 12, 10, 12]), [12, 12, 12, 12, 9, 12], [12, 12, 12, 12, 10, 12]])
        .setHorizontalAlignments(new Array(lastRow).fill(['left', 'left', 'center', 'center', 'right', 'left']))
        .setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK)

      sheets[s].getRangeList(['B2:C3', 'B5:C6', 'B8:C8', 'B10:C10']).setBackground('white').setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK)
      sheets[s].getRange('D2:D10').uncheck()
    }
    else if (sheetName === 'Export')
    {
      range = sheets[s].setColumnWidth(1, 17).setColumnWidth(2, 126).setColumnWidth(3, 70).setColumnWidths(4, 3, 100)
        .setColumnWidth(7, 75).setColumnWidths(8, 2, 25).setColumnWidth(10, 100).getDataRange();
      lastRow = range.getLastRow()
      lastCol = range.getLastColumn()

      range.setFontColor('black').setFontFamily('Arial').setFontSize(10).setNumberFormat('@').setVerticalAlignment('middle')
    }
    else if (sheetName === 'Last_Import' || sheetName === 'All_Active_Orders')
    {
      range = sheets[s].hideSheet().getDataRange();
      lastRow = range.getLastRow()
      lastCol = range.getLastColumn()
      sheets[s].setFrozenRows(1)

      range.setBackgrounds([new Array(lastCol).fill('#c9daf8'), ...new Array(lastRow - 1).fill(new Array(lastCol).fill('white'))]).setFontColor('black')
        .setFontFamily('Arial').setFontSize(10).setNumberFormat('@').setVerticalAlignment('middle')
    }
    else if (sheetName === 'Status')
      sheets[s].hideSheet().getDataRange().setBackground('white').setFontColor('black').setFontFamily('Arial').setFontSize(10)
        .setHorizontalAlignment('left').setNumberFormat('@').setVerticalAlignment('middle')
    else if (sheetName === 'Complete')
      continue;
  }

  spreadsheet.toast('All sheets were formatted.')
}

/**
 * This function formats all of the header information for the packing slip. It is intended to run after user has completed an order from the Packing Slip page.
 * 
 * @param {Sheet} sheet : The packing slip sheet
 * @param {Number} shippingAmount : The value of shipping
 * @author Jarren Ralf
 */
function applyFormattingToPackingSlip(sheet, spreadsheet, shippingAmount)
{
  const values = sheet.getSheetValues(1, 2, 14, 8);
  const ordNumber      = values[ 0][7]
  const subtotalAmount = values[ 2][7]
  const customerName   = values[ 7][0]
  const shippingMethod = values[13][0]
  const col = 9; // Number of columns on the packing slip
  const numItemsOnPageOne = 32;
  const boldTextStyle = SpreadsheetApp.newTextStyle().setBold(true).setFontSize(12).build();
  const pntLogo = SpreadsheetApp.newCellImage().toBuilder().setSourceUrl('http://cdn.shopify.com/s/files/1/0018/7079/0771/files/logoh_180x@2x.png?v=1613694206').build();
  const pntAddress = SpreadsheetApp.newRichTextValue().setText('3731 Moncton Street, Richmond, BC, V7E 3A5\nPhone: (604) 274-7238 Toll Free: (800) 895-4327\nwww.pacificnetandtwine.com')
    .setLinkUrl(91, 117, 'https://www.pacificnetandtwine.com/').build()
  const shipDate = SpreadsheetApp.newRichTextValue().setText('Ship Date: ' + Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), "dd MMMM yyyy"))
    .setTextStyle(0, 10, boldTextStyle).build();
  const emailHyperLink = SpreadsheetApp.newRichTextValue().setText('If you have any questions, please send an email to: websales@pacificnetandtwine.com')
        .setLinkUrl(52, 83, 'mailto:websales@pacificnetandtwine.com?subject=RE: [Pacific Net %26 Twine Ltd] ' + ordNumber + ' placed by ' + customerName).build()

  sheet.setColumnWidth(1, 40).setColumnWidth(2, 133).setColumnWidth(3, 162).setColumnWidth(4, 156)
    .setColumnWidth(5, 40).setColumnWidth(6, 95).setColumnWidth(7, 50).setColumnWidths(8, 2, 75)
    .setRowHeight(1, 20).setRowHeight(2, 40).setRowHeights(3, 4, 20).setRowHeight(7, 10).setRowHeights(8, 5, 20).setRowHeight(13, 10)
    .setRowHeight(14, 20).setRowHeight(15, 10).setRowHeights(16, 33, 20).setRowHeight(49, 10).setRowHeight(50, 20)
  .getRange(1, 1) // PNT Logo
    .setValue(pntLogo)
  .offset(0, 7, 6, 2) // The invoice header values on page one
    .setFontColor('black').setFontFamily('Arial').setFontStyle('normal').setBackground('white')
    .setFontSizes([[10, 10], [10, 9], [10, 10], [10, 10], [10, 10], [10, 10]])
    .setVerticalAlignments([['middle', 'middle'], ['top', 'top'], ['middle', 'middle'], ['middle', 'middle'], ['middle', 'middle'], ['middle', 'middle']])
    .setNumberFormats([['@', '@'], ['@', 'dd MMM yyyy'], ['@', '$#,##0.00'], ['@', '$#,##0.00'], ['@', '$#,##0.00'], ['@', '$#,##0.00']])
    .setHorizontalAlignments([['right', 'center'], ['right', 'center'], ['right', 'right'], ['right', 'right'], ['right', 'right'], ['right', 'right']])
    .setFontWeights(new Array(6).fill(['bold', 'normal']))
    .setValues([['Web Order Number', '=Last_Import!A2'], ['Ordered Date', '=INDEX(SPLIT(Last_Import!P2, " "), 1, 1)'], 
      ['Subtotal Amount:', subtotalAmount], ['Shipping Amount:', shippingAmount], 
      ['Taxes:', '=Items_Tax+Shipping_Tax'], ['Order Total:', '=SUM(Order_Subtotals)']])
  .offset(3, -7) // PNT Address
    .setRichTextValue(pntAddress)
  .offset(4, 0, 5) // The value "SHIP" in the header of the packing slip
    .setBackground('#d9d9d9').setBorder(true, true, true, true, false, false).setFontColor('black').setFontFamily('Arial')
    .setFontLine('none').setFontSize(14).setFontStyle('normal').setFontWeight('bold').setHorizontalAlignment('center').setNumberFormat('@')
    .setVerticalAlignment('middle').setVerticalText(true).setValue('SHIP')
  .offset(0, 1, 5, 2) // The "SHIP" values
    .mergeAcross().setBackground('white').setBorder(true, true, true, true, false, false).setFontColor('black').setFontFamily('Arial')
    .setFontLine('none').setFontSize(10).setFontStyle('normal').setFontWeight('normal').setHorizontalAlignment('left').setNumberFormat('@')
    .setVerticalAlignment('middle').setVerticalText(false).setFormulas([['Last_Import!AI2', null], ['Last_Import!AM2', null], 
      ['Last_Import!AK2', null], ['Last_Import!AN2&", "&Last_Import!AP2&", "&Last_Import!AO2&", "&Last_Import!AQ2', null], ['Last_Import!AR2', null]])
  .offset(0, 3, 5, 1) // The value "BILL" in the header of the packing slip
    .setBackground('#d9d9d9').setBorder(true, true, true, true, false, false).setFontColor('black').setFontFamily('Arial')
    .setFontLine('none').setFontSize(14).setFontStyle('normal').setFontWeight('bold').setHorizontalAlignment('center').setNumberFormat('@')
    .setVerticalAlignment('middle').setVerticalText(true).setValue('BILL')
  .offset(0, 1, 5, 4) // The "BILL" values
    .mergeAcross().setBackground('white').setBorder(true, true, true, true, false, false).setFontColor('black').setFontFamily('Arial')
    .setFontLine('none').setFontSize(10).setFontStyle('normal').setFontWeight('normal').setHorizontalAlignment('left').setNumberFormat('@')
    .setVerticalAlignment('middle').setVerticalText(false).setFormulas([['Last_Import!Y2', null, null, null], ['Last_Import!AC2', null, null, null], ['Last_Import!AA2', null, null, null], 
      ['Last_Import!AD2&", "&Last_Import!AF2&", "&Last_Import!AE2&", "&Last_Import!AG2', null, null, null], ['Last_Import!AH2', null, null, null]])
  .offset(6, -5, 1, col) // The shipping values
    .setBackground('white').setBorder(true, true, true, true, false, false).setFontColor('black').setFontFamily('Arial')
    .setFontLine('none').setFontStyle('normal').setNumberFormat('@').setVerticalAlignment('middle').setVerticalText(false)
    .setFontSizes([[12, ...new Array(col - 1).fill(10)]])
    .setFontWeights([['bold', ...new Array(col - 1).fill('normal')]])
    .setHorizontalAlignments([[...new Array(col - 3).fill('left'), 'right', 'left', 'left']])
    .setValues([['Via', shippingMethod, ...new Array(col - 2).fill('')]])
  .offset(0, 1, 1, 2).merge() // Shipping method cells
  .offset(0, 5, 1, 3).merge() // Ship Date cells
    .setRichTextValue(shipDate)
  .offset(numItemsOnPageOne + 4, -6, 1, 5) // The hyperlinked email at the bottom of the page
    .merge().setRichTextValue(emailHyperLink)
}

/**
 * This function takes the given string, splits it at the chosen delimiter, and capitalizes the first letter of each perceived word.
 * 
 * @param {String} str : The given string
 * @param {String} delimiter : The delimiter that determines where to split the given string
 * @return {String} The output string with proper case
 * @author Jarren Ralf
 */
function capitalizeSubstrings(str, delimiter)
{
  var numLetters;
  var words = str.toString().split(delimiter); // Split the string at the chosen location/s based on the delimiter

  for (var word = 0, string = ''; word < words.length; word++) // Loop through all of the words in order to build the new string
  {
    numLetters = words[word].length;

    if (numLetters == 0) // The "word" is a blank string (a sentence contained 2 spaces)
      continue; // Skip this iterate
    else if (numLetters == 1) // Single character word
    {
      if (words[word][0] !== words[word][0].toUpperCase()) // If the single letter is not capitalized
        words[word] = words[word][0].toUpperCase(); // Then capitalize it
    }
    else if (numLetters == 2 && words[word].toUpperCase() === 'PO') // So that PO Box is displayed correctly
      words[word] = words[word].toUpperCase();
    else
    {
      /* If the first letter is not upper case or the second letter is not lower case, then
       * capitalize the first letter and make the rest of the word lower case.
       */
      if (words[word][0] !== words[word][0].toUpperCase() || words[word][1] !== words[word][1].toLowerCase())
        words[word] = words[word][0].toUpperCase() + words[word].substring(1).toLowerCase();
    }

    string += words[word] + delimiter; // Add a blank space at the end
  }

  string = string.slice(0, -1); // Remove the last space

  return string;
}

/**
 * This function...
 * 
 * @author Jarren Ralf
 */
function changeShippingAddress()
{
  var html = HtmlService.createHtmlOutputFromFile('changePickUpAddress.html')
    .setWidth(400)
    .setHeight(250);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showModalDialog(html, "Change Ship-To Address");
}

/**
 * This function...
 * 
 * @author Jarren Ralf
 */
function clearExportPage()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getActiveSheet();

  if (sheet.getSheetName() !== 'Export')
  {
    spreadsheet.toast('You must be on the Export sheet to run this function. Please try again.')
    spreadsheet.getSheetByName('Export').activate().clear();
  }
  else
    sheet.clear(); 
}

/**
 * This function creates the triggers associated with this spreadsheet.
 * 
 * @author Jarren Ralf
 */
function createTrigger()
{
  ScriptApp.newTrigger('onChange').forSpreadsheet('1QIKO0KcWPYoP4yR5c22jvbY0Ldsx9tO4MqyrMiTTzBc').onChange().create() // This is an installable onChange trigger
}

/**
 * This function displays the shipping calculator as a sidebar.
 * 
 * @author Jarren Ralf
 */
function displayShippingCalculator()
{
  var sidebar = HtmlService.createHtmlOutputFromFile("shippingCalculator.html");
  sidebar.setTitle("Calculator").setWidth(600);
  SpreadsheetApp.getUi().showSidebar(sidebar);
}

/**
 * This function deletes the two pdf files that the user has created in the google drive.
 * 
 * @param {String} id1 : The file id of one of the files
 * @param {String} id2 : The file id of the other file
 * @author Jarren Ralf
 */
function deleteFiles(id1, id2)
{
  DriveApp.getFileById(id1).setTrashed(true)
  DriveApp.getFileById(id2).setTrashed(true)
}

/**
 * This function deletes of the Triggers associated with the user in regards to this project.
 * 
 * @author Jarren Ralf
 */
function deleteTriggers()
{
  ScriptApp.getProjectTriggers().map(trigger => ScriptApp.deleteTrigger(trigger));
}

/**
 * This function checks if the given phone number has a leading 1 or not.
 * 
 * @param {String} p : The phone number of a customer
 * @return {Boolean} Whether the phone number contains 11 characters and starts with a 1 or not.
 * @author Jarren Ralf
 */
function doesPhoneNumberStartWithOne(p)
{
  return p.length === 11 && p[0] === '1';
}

/**
 * This function launches a modal dialog box which allows the user to click a download button, which will lead to 
 * two pdf files being downloaded.
 * 
 * @author Jarren Ralf
 */
function downloadButton()
{
  const spreadsheet = SpreadsheetApp.getActive()
  const invoiceSheet = spreadsheet.getSheetByName('Invoice')
  const packingSlipSheet = spreadsheet.getSheetByName('Packing Slip')
  const customerName = invoiceSheet.getSheetValues(8, 6, 1, 1)[0][0]
  const invoicePdf = getAsBlob(spreadsheet, invoiceSheet).getAs('application/pdf').setName(customerName + "_Invoice.pdf")
  const packingSlipPdf = getAsBlob(spreadsheet, packingSlipSheet).getAs('application/pdf').setName(customerName + "_PackingSlip.pdf")
  var htmlTemplate = HtmlService.createTemplateFromFile('DownloadButton');
  const invoiceFile = DriveApp.createFile(invoicePdf)
  const packingSlipFile = DriveApp.createFile(packingSlipPdf)
  htmlTemplate.url1 = invoiceFile.getDownloadUrl();
  htmlTemplate.url2 = packingSlipFile.getDownloadUrl();
  htmlTemplate.fileId1 = invoiceFile.getId();
  htmlTemplate.fileId2 = packingSlipFile.getId();
  var html = htmlTemplate.evaluate().setWidth(250).setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(html, 'Export');
}

/**
 * This function sends the customer an email with an attached copy of the packing slip as a pdf. The pdf and email both contain hyperlinks for the
 * tracking number (if applicable) and the reply email.
 * 
 * @param {Boolean} comments : A boolean for whether the user wants to have additional comments in the body of the email
 * @author Jarren Ralf
 */
function emailPackingSlip(comments)
{
  const spreadsheet = SpreadsheetApp.getActive();
  const packingSlip = spreadsheet.getSheetByName('Invoice').activate();
  const packingSlipValues = packingSlip.getSheetValues(1, 2, 14, 8);

  if (packingSlipValues[13][0] !== 'Select Shipping Method')
  {
    const recipientEmail = spreadsheet.getSheetByName('Last_Import').getSheetValues(2, 2, 1, 1)[0][0];
    //const recipientEmail = 'adrian@pacificnetandtwine.com' // For testing  
    const orderNumber = packingSlipValues[0][7];
    const billingName = packingSlipValues[7][4];

    // Read in and set the appropriate variables on the html template
    if (isBlank(packingSlipValues[13][2]))
      var templateHtml = (comments) ? HtmlService.createTemplateFromFile('customEmail') : HtmlService.createTemplateFromFile('email');
    else // Tracking number is included
    {
      var linkUrl = "";

      // Check the shipping method and make the relevant changes
      switch (packingSlipValues[13][0])
      {
        case 'Post Tracked Packet':
        case 'Post Expedited Parcel':
        case 'Post Xpress Post':
          linkUrl += "https://www.canadapost.ca/track-reperage/en#/search?searchFor=" + packingSlipValues[13][2];
          break;
        case 'Purolator Ground':
        case 'Purolator Express':
          linkUrl += "https://www.purolator.com/purolator/ship-track/tracking-details.page?pin=" + packingSlipValues[13][2];
          break;
        case 'UPS Standard':
        case 'UPS Express':
          linkUrl += "https://www.ups.com/track?loc=en_CA&tracknum=" + packingSlipValues[13][2];
          break;
      }

      const hyperlink = SpreadsheetApp.newTextStyle().setFontSize(11).setUnderline(true).setForegroundColor('#1155cc').build();
      const hyperlinkRichText = SpreadsheetApp.newRichTextValue().setText(packingSlipValues[13][2]).setTextStyle(hyperlink).setLinkUrl(linkUrl).build();
      packingSlip.getRange(14, 4).setRichTextValue(hyperlinkRichText)

      var templateHtml = (comments) ? HtmlService.createTemplateFromFile('customEmailWithTrackingNumber') : HtmlService.createTemplateFromFile('emailWithTrackingNumber');
      templateHtml.trackingNumber = packingSlipValues[13][2];
      templateHtml.url = linkUrl
    }
    
    templateHtml.recipientName = billingName.split(' ', 1)[0];
    templateHtml.shippingMethod = packingSlipValues[13][0];

    if (comments)
      templateHtml.comments = prompt('Please type you email comments:')

    var emailSubject = 'RE: [Pacific Net & Twine Ltd] ' + orderNumber + ' placed by ' + billingName;
    var emailSignature = '<p>If you have any questions, please click reply or send an email to: <a href="mailto:websales@pacificnetandtwine.com?subject=RE: [Pacific Net %26 Twine Ltd] ' + 
      orderNumber + ' placed by ' + billingName + '">websales@pacificnetandtwine.com</a></p>'
    var message = templateHtml.evaluate().append(emailSignature).getContent(); // Get the contents of the html document
    var packingSlipPDF = getAsBlob(spreadsheet, packingSlip).getAs('application/pdf').setName(orderNumber + ".pdf")

    // Fire an email with following chosen parameters
    GmailApp.sendEmail(recipientEmail, 
                        emailSubject, 
                        '',
                      {   replyTo: 'websales@pacificnetandtwine.com',
                             //from: 'pntnoreply@gmail.com',
                             name: 'PNT Web Sales',
                         htmlBody: message,
                      attachments: packingSlipPDF});

    spreadsheet.toast('Email Successfully Sent to ' + billingName);
  }
  else
    Browser.msgBox('Please select a shipping method.')
}

/**
 * This function sends the customer an email with an attached copy of the packing slip as a pdf. The pdf and email both contain hyperlinks for the
 * tracking number (if applicable) and the reply email. Running this function launchs a dialogue box to give the user an option to write a custom message
 * to the recipient.
 * 
 * @author Jarren Ralf
 */
function emailPackingSlipWithComments()
{
  emailPackingSlip(true)
}

/**
 * This function takes the shopify data on the import page and prepares it for export into Adagio. 
 * This function is run either via an onEdit event where a copy-paste of many columns of data has 
 * occured, or it can be run manually.
 * 
 * @param {Object[][]}  importData  : The import data
 * @param  {Sheet}      exportSheet : The export sheet 
 * @param {Spreadsheet} spreadsheet : The active spreadsheet
 * @param {Number}   shippingAmount : The cost of shipping
 * @param {Object[][]}  itemValues  : The item values that come from the packing slip
 * @author Jarren Ralf
 */
function exportData(importData, exportSheet, spreadsheet, shippingAmount, itemValues)
{
  if (arguments.length === 0) // If the function is being run manually (no arguments are passed) 
  {
    exportSheet = spreadsheet.getSheetByName('Export');
    importData = SpreadsheetApp.getActiveSheet().getDataRange().getValues()
  }

  const nCols = importData[0].length;
  const colours = [new Array(nCols).fill('#c9daf8')]; // To highlight the import data alternating by order #
  const numCols = 10; // Number of columns in the export data
  var headerLine, shippingLine, productLine, country, province, numOrders = 0, shippingCosts = [], exportData = [], isFreightLineIncluded = [], backgrounds = [];
  
  for (var i = 1; i < importData.length; i++) // Loop through each row of the import data (skip the header)
  {
    if (isNewOrder(importData, i)) // The current row is a new order (the first order is considered a new order)
    {
      /* This code block checks if a freight line needs to be included in the previous order.
       * Note: The first order does not have a previous order.
       */
      if (isNotFirstOrder(i) && isFreightLineIncluded[numOrders - 1])
      {
        // The shipping cost calculation might yield a negative value if the shipped quantities were changed in Shopify
        if (shippingCosts[numOrders - 1] < 0)
        { 
          // Highlight the negative value in red and set the freight to 0
          shippingCosts[numOrders - 1] = 0;
          backgrounds.push(['white', 'white', '#f4cccc', ...new Array(numCols - 3).fill('white')])
        }
        else
          backgrounds.push(new Array(numCols).fill('white'))

        exportData.push(getFreightLine(shippingCosts[numOrders - 1], country)); // Add a freight line
      }

      country = (isNotBlank(importData[i][42])) ? importData[i][42].toUpperCase() : importData[i][32].toUpperCase();
      province = (isNotBlank(importData[i][41])) ? importData[i][41].toUpperCase() : importData[i][31].toUpperCase();

      // The shipping charges for USA orders don't contain GST and PST, but the canadian orders do
      (country !== 'CA') ? shippingCosts.push(importData[i][9] - importData[i][51]) : shippingCosts.push(importData[i][9] - importData[i][51]/1.12);

        headerLine =   getHeaderLine(importData[i], province, country);
      shippingLine = getShippingLine(importData[i]);
       productLine =  getProductLine(importData[i], province, country);
       exportData.push(headerLine, shippingLine, productLine);
      backgrounds.push(new Array(numCols).fill('#c9daf8'), new Array(numCols).fill('#e0e9f9'), new Array(numCols).fill('white'))

      // Put a true/false at the front of the array depending on whether the order is a pick up or not
      isFreightLineIncluded.push(!isPickUp(importData[i]));

      numOrders++; // Count the number of orders
    }
    // Add an additional product line if the SKU is not blank and the item is fulfilled
    else if (isSKU_NotBlank(importData[i]) && isItemFulfilled(importData[i], !isFreightLineIncluded[numOrders - 1]))
    {
      exportData.push(getProductLine(importData[i], province, country));
      backgrounds.push(new Array(numCols).fill('white'));
    }
    else if (isLineItemPriceNonZero(importData[i]) && isShippingRequiredForLineItem(importData[i]))
      // Add the additional freight cost to the default freight charge (which is usually $30)
      shippingCosts[numOrders - 1] += importData[i][18]; 

    (numOrders%2 === 0) ? colours.push(new Array(nCols).fill('#e0e9f9')) : colours.push(new Array(nCols).fill('white'))
  }

  if (arguments.length > 3) // Generate the 'Detail' line items with the packing slip data rather than the shopify data
  {
    exportData.pop()
    backgrounds.pop()
    var qty;

    itemValues.map(val => {
      if (isNotBlank(val[1]))
      {
        qty = val[0].split(' ', 1)[0];
        backgrounds.push(new Array(numCols).fill('white'))

        exportData.push(['D', 
          val[1], // sku
          val[7], // price
          qty, 
          qty, 
          getTaxCode_LineItem(importData[1], province, country), // Set the tax code
          ...new Array(4).fill(null)])
        }
    })

    shippingCosts[0] = shippingAmount;
  }

  if (isFreightLineIncluded[numOrders - 1])
  {
    // The shipping cost calculation might yield a negative value if the shipped quantities were changed in Shopify
    if (shippingCosts[numOrders - 1] < 0)
    { // Highlight the negative value in red
      shippingCosts[numOrders - 1] = 0;
      backgrounds.push(['white', 'white', '#f4cccc', ...new Array(numCols - 3).fill('white')]);
    }
    else
      backgrounds.push(new Array(numCols).fill('white'))

    exportData.push(getFreightLine(shippingCosts[numOrders - 1], country)); // Add a FREIGHT line to the end of the export data
  }
  
  if (exportData.length !== 0)
    exportSheet.getRange(exportSheet.getLastRow() + 1, 1, exportData.length, numCols).setBackgrounds(backgrounds).setValues(exportData);
}

/**
 * This function reformats a valid phone number into (###) ###-####, unless there are too many/few digits in the number, in which case the original string is returned.
 * It handles inputs that include leading ones and pluses, as well as strings that contain or don't contain parenthesis.  
 * 
 * @param {Number} num : The given phone number
 * @return Returns a reformatted phone number
 * @author Jarren Ralf
 */
function formatPhoneNumber(num)
{
  var ph = num.toString().trim().replace(/['\])}[\s{(+-]/g, ''); // Remove any brackets, braces, parenthesis, apostrophes, dashes, plus symbols, and blank spaces

  return (ph.length === 10 && ph[0] !== '1') ? '(' + ph.substring(0, 3) + ') ' + ph.substring(3, 6) + '-' + ph.substring(6) : 
         (ph.length === 11 && ph[0] === '1') ? '(' + ph.substring(1, 4) + ') ' + ph.substring(4, 7) + '-' + ph.substring(7) : num;
}

/**
 * This function reformats a valid canadian postal code into A1A 1A1, unless there are too many/few digits in the number, in which case the original string is returned.
 * 
 * @param {Number} num : The given postal code
 * @return Returns a reformatted candian postal code
 * @author Jarren Ralf
 */
function formatPostalCode(num)
{
  var postCode = num.toString().trim().toUpperCase(); 

  return (postCode.length === 6) ? postCode.substring(0, 3) + ' ' + postCode.substring(3, 6) : postCode;
}

/**
 * This function converts the given sheet into a BLOB object. Based on the second argument, namely which sheet is getting converted, certain parameters are 
 * set that lead to the BLOB object being stored as a csv or pdf file.
 * 
 * @license MIT
 * 
 * © 2020 xfanatical.com. All Rights Reserved.
 * @param {Spreadsheet} spreadsheet : The specific spreadsheet that will be used to convert the export page into a BLOB (Binary Large Object)
 * @return The packing slip sheet as a BLOB object that will eventually get converted to pdf document that will be attached to an email sent to the customer
 * @author Jason Huang
 */
function getAsBlob(spreadsheet, sheet)
{
  // A credit to https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
  // these parameters are reverse-engineered (not officially documented by Google)
  // they may break overtime.
  var exportUrl = spreadsheet.getUrl().replace(/\/edit.*$/, '') + '/export?'
      + 'exportFormat=pdf'
      + '&format=pdf'
      + '&size=LETTER'
      + '&portrait=true'
      + '&fitw=true&top_margin=0.25&bottom_margin=0.25&left_margin=0.25&right_margin=0.25'           
      + '&sheetnames=false&printtitle=false&pagenum=UNDEFINED'
      + '&gridlines=false&fzr=FALSE'
      + '&gid=' + sheet.getSheetId();

  var response;

  for (var i = 0; i < 5; i++)
  {
    response = UrlFetchApp.fetch(exportUrl, {
      muteHttpExceptions: true,
      headers: { 
        Authorization: 'Bearer ' +  ScriptApp.getOAuthToken(),
      },
    })

    if (response.getResponseCode() === 429)
      Utilities.sleep(3000) // printing too fast, retrying
    else
      break;
  }
  
  if (i === 5)
    throw new Error('Printing failed. Too many sheets to print.');

  return response.getBlob()
}

/**
 * This function returns the freight line for the export data. The cost of freight is rounded to 2 decimals.
 * If the country is USA the tax code is set to 4, otherwise it is set to 0. 
 * 
 * @param {Object[]} cost  : The cost of shipping
 * @param {String} country : The abbreviation for the country
 * @return {Object[]} The freight line for the export data
 * @author Jarren Ralf
 */
function getFreightLine(cost, country)
{
  return ['D', 'FREIGHT', twoDecimals(cost), 1, 1, (country !== 'CA') ? 4 : 0, ...new Array(4).fill(null)];
}

/**
 * This function takes the import data for the given row and creates a header line for the export data.
 * It sets whether the order is a pick up or will be shipped, sets the inventory location, and gets the 
 * tax code for the order.
 * 
 * @param {Object[]} data : The given row of import data
 * @param {String} province : The abbreviation for the province (or state)
 * @param {String} country  : The abbreviation for the country
 * @return {Object[]} The header line for the export data
 * @author Jarren Ralf
 */
function getHeaderLine(data, province, country)
{
  return ['H',                                 // Line Type Code (H: Header, S: Shipping, D: Detail)
    "C4458",                                   // Account Number
    'Ord# ' + data[0].replace(/\D/g, ''),      // Order Number
    (isPickUp(data)) ? 'PICK UP' : 'MAIL',     // Shipping Method
    getTaxCode_Order(data, province, country), // Tax Code
    (isNotBlank(data[41])) ? data[41].toUpperCase() : data[31].toUpperCase(), // Province / State
    getInventoryLocation(data),                // The inventory location
    ...new Array(3).fill(null)];
}

/**
 * This function takes the import data for the given row and assigns the location that the inventory will
 * be pulled from.
 * 
 * @param {Object[]} data : The given row of import data
 * @return {Number} The inventory location: 100 (Default), 200, or 300
 * @author Jarren Ralf
 */
function getInventoryLocation(data)
{
  return (data[14] === "Parksville") ? 200 : (data[14] === "Prince Rupert") ? 300 : 100;
}

/**
 * This function takes the import data for the given row and creates a product line for the export data.
 * The unit price is set for the product, including any discounts and then rounded to 2 decimals.
 * 
 * @param {Object[]} data   : The given row of import data
 * @param {String} province : The abbreviation for the province (or state)
 * @param {String} country  : The abbreviation for the country
 * @return {Object} The product line for the export data
 * @author Jarren Ralf
 */
function getProductLine(data, province, country)
{
  return ['D', // Line Type Code (H: Header, S: Shipping, D: Detail)
    '\'' +  data[20],  // SKU (the leading apostrophe is placed infront of the sku in order to eliminate the conversion of the perceived number to scientific notation when imported into excel)
            twoDecimals(data[18] - data[59]/data[16]), // Discounted Unit Price
            data[16],  // Quantity
            data[16],  // Quantity
            getTaxCode_LineItem(data, province, country), // Set the tax code
    ...new Array(4).fill(null)];
}

/**
 * This function takes the import data for the given row and creates a shipping line for the export data.
 * 
 * @param {Object[]} data : The given row of import data
 * @return {Object} The shipping line for the export data
 * @author Jarren Ralf
 */
function getShippingLine(data)
{
  var phoneNumber = (isNotBlank(data[43])) ? data[43].toString().replace(/\D/g,'') : data[33].toString().replace(/\D/g,'');

  if (doesPhoneNumberStartWithOne(phoneNumber))
    phoneNumber = phoneNumber.substring(1); // Remove the leading 1 in the phone number

  return shippingLine = ['S', // Line Type Code (H: Header, S: Shipping, D: Detail)
    (isNotBlank(data[34])) ? toProper(data[34]) : toProper(data[24]),                 // Name
    (isNotBlank(data[38])) ? toProper(data[38]) : toProper(data[28]),                 // Company Name
    (isNotBlank(data[36])) ? toProper(data[36]) : toProper(data[26]),                 // Address 1
    (isNotBlank(data[37])) ? toProper(data[37]) : toProper(data[27]),                 // Address 2
    (isNotBlank(data[39])) ? toProper(data[39]) : toProper(data[29]),                 // City
    (isNotBlank(data[40])) ? formatPostalCode(data[40]) : formatPostalCode(data[30]), // Postal Code / Zip Code
    (isNotBlank(data[41])) ? data[41].toUpperCase() : data[31].toUpperCase(),         // Province / State
    (isNotBlank(data[42])) ? data[42].toUpperCase() : data[32].toUpperCase(),         // Country
    phoneNumber];                                                                     // Phone Number
}

/**
 * This function takes the given row of import data and determines the tax code based on the 
 * shipping method, billing province/country, and shipping province/country for each line item.
 * 
 * @param {String[]} data : The given row of import data
 * @param {String} province : The abbreviation for the province (or state)
 * @param {String} country  : The abbreviation for the country
 * @return {String} The tax code
 * @author Jarren Ralf
 */
function getTaxCode_LineItem(data, province, country)
{
  // Is the shipping method a PICK UP or destination one of the following: British Columbia or Nova Scotia
  if (isPickUp(data) || province === 'BC' || province === 'NS' || province === 'SK' || province === 'QC' || province === 'PE' || province === 'NB' || province === 'NL')
    return 0;
  else if (country !== 'CA') // Country is not Canada
    return 4;
  else if (province === 'AB' || province === 'NT' || province === 'NU' || province === 'YT'|| province === 'MB'|| province === 'ON')
    return 2;
  else
    return '';
}

/**
 * This function takes the given row of import data and determines the tax code based on the 
 * shipping method, billing province/country, and shipping province/country for each order.
 * 
 * @param {String[]} data   : The given row of import data
 * @param {String} province : The abbreviation for the province (or state)
 * @param {String} country  : The abbreviation for the country
 * @return {String} The tax code
 * @author Jarren Ralf
 */
function getTaxCode_Order(data, province, country)
{
  // Is the shipping method a PICK UP or destination one of the following: British Columbia or Manitoba
  if (isPickUp(data) || province === 'BC' || province === 'MB')
    return 'BC';
  else if (province === 'AB' || province === 'QC' || province === 'YT' || province === 'NT' || province === 'NU' || province === 'SK')
    return 'EXTPRO';
  else if (country !== 'CA') // Country is not Canada
    return 'EXPORT';
  else if (province === 'ON' || province === 'NL' || province === 'NB' || province === 'PE') // Ontario or Newfoundland or New Brunswick
    return 'HSTPRO';
  else if (province === 'NS') // Nova Scotia
    return 'HSTCF2';
  else
    return 'No Tax Code';
}

/**
 * This function checks if the given string is blank or not.
 * 
 * @param {String} string : The given string.
 * @return Returns a true boolean if the given string is blank, false otherwise.
 * @author Jarren Ralf
 */
function isBlank(string)
{
  return string === '';
}

/**
 * This function checks if a particular line item is fulfilled or not. Technically "Pick Up" 
 * orders are unfulfilled, but return true anyways
 * 
 * @param  {Object[][]} values : The import data
 * @param  {Boolean} isOrderInStorePickUp : Whether the order is a pick up or not
 * @return {Boolean} Whether the given item is considered fulfilled or not.
 * @author Jarren Ralf
 */
function isItemFulfilled(data, isOrderInStorePickUp)
{
  return (isOrderInStorePickUp) ? true : data[23] === 'fulfilled';
}

/**
 * This function checks if the LineItem price is not zero
 * 
 * @param {String[]} data : The given row of import data
 * @return {Boolean} Whether the LineItem prices is non zero.
 * @author Jarren Ralf
 */
function isLineItemPriceNonZero(data)
{
  return data[18] !== 0;
}

/**
 * This function checks if the given row of import data is a pick up or not
 * 
 * @param {String[]} data : The given row of import data
 * @return {Boolean} Whether the Shipping Method parameter says either Richmond, Parksville, or Prince Rupert.
 * @author Jarren Ralf
 */
function isPickUp(data)
{
  return data[14] == "Richmond" || data[14] == "Parksville" || data[14] == "Prince Rupert";
}

/**
 * This function checks if the item at the current line requires shipping 
 * (FREIGHT - Doesn't require shipping).
 * 
 * @param {String[]} data : The given row of import data
 * @return {Boolean} Whether the 'Lineitem requires shipping' field is FALSE or not.
 * @author Jarren Ralf
 */
function isShippingRequiredForLineItem(data)
{
  return data[21] === false;
}

/**
 * This function checks if the current row of import data is a new order or part of the previous one. 
 * The first order is considered a new order.
 * 
 * @param {Object[][]} data : The import data
 * @param   {String}    i   : The current row number of the import data
 * @return {Boolean} Whether the current row Name parameter matches the previous order (row above the current) Name parameter or not
 * @author Jarren Ralf
 */
function isNewOrder(data, i)
{
  return data[i][0] !== data[i - 1][0] && isNotBlank(data[i][0]);
}

/**
 * This function checks if the given string is not blank.
 * 
 * @param {String} str : The given string
 * @return {Boolean} Whether the given string is not blank or it is blank.
 * @author Jarren Ralf
 */
function isNotBlank(str)
{
  return str !== '';
}

/**
 * This function checks if the given row number of the import data is not the first order or it is.
 * 
 * @param {Number} i : The row number of the current order
 * @return {Boolean} Whether the row number is not line 1 or it is.
 * @author Jarren Ralf
 */
function isNotFirstOrder(i)
{
  return i !== 1;
}

/**
 * This function checks if the SKU is not blank.
 * 
 * @param {String[]} data : The given row of import data
 * @return {Boolean} Whether the SKU is not blank or if it is blank.
 * @author Jarren Ralf
 */
function isSKU_NotBlank(data)
{
  return isNotBlank(data[20]);
}

/**
 * This function ...
 * 
 * @author Jarren Ralf
 */
function packingSlip_BackOrder()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getActiveSheet();

  if (sheet.getSheetName() !== 'Invoice' && sheet.getSheetName() !== 'Packing Slip')
  {
    spreadsheet.toast('You must be on the Invoice sheet to run this function. Please try again.')
    spreadsheet.getSheetByName('Invoice').activate();
  }
  else
  {
    const currentOrder = sheet.getSheetValues(1, 9, 1, 1)[0][0]
    const statusPage = spreadsheet.getSheetByName('Status Page');
    const range = statusPage.getRange(3, 1, statusPage.getLastRow() - 2, 4);
    var values = range.getValues();

    for (var i = 0; i < values.length; i++)
    {
      if (values[i][0] === currentOrder)
      {
        values[i][3] = 'Items Backordered, Customer will wait.'
        break;
      }
    }

    range.setValues(values)

    var numPages = sheet.getSheetValues(42, 9, 1, 1)[0][0];

    if (isBlank(numPages))
      var itemValues_PackingSlip = sheet.getSheetValues(17, 1, 24, 8);
    else
    {
      numPages = numPages.split(' ')[3]
      var itemValues_PackingSlip = sheet.getSheetValues(17, 1, 24, 8);

      for (var p = 0; p < numPages; p++)
        itemValues_PackingSlip.push(...sheet.getSheetValues(52 + p*42, 1, 32, 8))
    }

    const shippingAmount = sheet.getSheetValues(4, 9, 1, 1)[0][0];
    const activeOrdersPage = spreadsheet.getSheetByName('All_Active_Orders')
    const activeOrdersValues = activeOrdersPage.getSheetValues(1, 1, activeOrdersPage.getLastRow(), activeOrdersPage.getLastColumn())
    const values_ExportPage = [activeOrdersValues[0]]; // The shopify data used to create the export data for Adagio; initialize with the header

    for (var j = 1; j < activeOrdersValues.length; j++)
    {
      if (activeOrdersValues[j][0] === currentOrder)
      {
        values_ExportPage.push(activeOrdersValues[j])
        break;
      }
    }

    const completedOrdersPage = spreadsheet.getSheetByName('Completed Orders');
    const lastRow = completedOrdersPage.getLastRow();
  
    const ordersOnCompletePage = (lastRow === 0) ? [[currentOrder + ' - Back Order']] : 
      Array.from(new Set(completedOrdersPage.getSheetValues(1, 1, lastRow, 1)
        .concat([[currentOrder + ' - Back Order']]).sort((a, b) => (a[0] < b[0]) ? 1 : -1).map(JSON.stringify)), JSON.parse)

    completedOrdersPage.getRange(1, 1, lastRow).setValues(ordersOnCompletePage)

    exportData(values_ExportPage, spreadsheet.getSheetByName('Export'), spreadsheet, shippingAmount, itemValues_PackingSlip);
    statusPage.activate()
    applyFormattingToPackingSlip(sheet, spreadsheet, shippingAmount)
  }
}

/**
 * This function ...
 * 
 * @author Jarren Ralf
 */
function packingSlip_Complete()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getActiveSheet();

  if (sheet.getSheetName() !== 'Invoice' && sheet.getSheetName() !== 'Packing Slip')
  {
    spreadsheet.toast('You must be on the Invoice sheet to run this function. Please try again.')
    spreadsheet.getSheetByName('Invoice').activate();
  }
  else
  {
    const currentOrder = sheet.getSheetValues(1, 9, 1, 1)[0][0]

    const statusPage = spreadsheet.getSheetByName('Status Page').activate();
    const numCols_StatusPage = statusPage.getLastColumn();
    const values_StatusPage = statusPage.getSheetValues(3, 1, statusPage.getLastRow() - 2, numCols_StatusPage);
    const remainingOrders_StatusPage = values_StatusPage.filter(v => v[0] != currentOrder);
    const numOrders = remainingOrders_StatusPage.length;
    const numberFormats = new Array(numOrders).fill(['@', '#', ...new Array(numCols_StatusPage - 3).fill('@'), "dd MMM yyyy"])

    statusPage.getRange(3, 1, numOrders, numCols_StatusPage).setNumberFormats(numberFormats).setValues(remainingOrders_StatusPage)
    statusPage.deleteRows(numOrders + 3, values_StatusPage.length - numOrders) // Delete the last rows

    var isCurrentOrder, isFirstRowOfCurrentData = true;
    
    const activeOrdersPage = spreadsheet.getSheetByName('All_Active_Orders')
    const numCols_OrdersPage = activeOrdersPage.getLastColumn();
    const values_OrdersPage = activeOrdersPage.getSheetValues(2, 1, activeOrdersPage.getLastRow() - 1, numCols_OrdersPage);
    const values_ExportPage = [activeOrdersPage.getSheetValues(1, 1, 1, numCols_OrdersPage)[0]]; // The shopify data used to create the export data for Adagio; initialize with the header

    const remainingOrders_OrdersPage = values_OrdersPage.filter(v => {
      isCurrentOrder = v[0] == currentOrder;

      if (isCurrentOrder && isFirstRowOfCurrentData)
      {
        values_ExportPage.push(v);
        isFirstRowOfCurrentData = false;
      }

      return !isCurrentOrder;
    });
    
    const numRows = remainingOrders_OrdersPage.length;

    activeOrdersPage.getRange(2, 1, numRows, numCols_OrdersPage).setValues(remainingOrders_OrdersPage)
    activeOrdersPage.deleteRows(numRows + 2, values_OrdersPage.length - numRows) // Delete the last rows

    var lastCol = 9;
    var numPages = sheet.getSheetValues(50, lastCol, 1, 1)[0][0];

    if (isBlank(numPages))
      var itemValues_PackingSlip = sheet.getSheetValues(17, 1, 32, lastCol - 1);
    else
    {
      numPages = numPages.split(' of ')[1]
      var itemValues_PackingSlip = sheet.getSheetValues(17, 1, 32, lastCol - 1);

      for (var p = 0; p < numPages; p++)
        itemValues_PackingSlip.push(...sheet.getSheetValues(59 + p*49, 1, 39, lastCol - 1))
    }

    const shippingAmount = sheet.getSheetValues(4, lastCol, 1, 1)[0][0];

    const completedOrdersPage = spreadsheet.getSheetByName('Completed Orders');
    const lastRow = completedOrdersPage.getLastRow();
  
    const ordersOnCompletePage = (lastRow === 0) ? [[currentOrder + ' - Complete']] : 
      completedOrdersPage.getSheetValues(1, 1, completedOrdersPage.getLastRow(), 1).concat([[currentOrder + ' - Complete']]).sort((a, b) => (a[0] < b[0]) ? 1 : -1);

    completedOrdersPage.getRange(1, 1, ordersOnCompletePage.length).setValues(ordersOnCompletePage)

    exportData(values_ExportPage, spreadsheet.getSheetByName('Export'), spreadsheet, shippingAmount, itemValues_PackingSlip);
    applyFormattingToPackingSlip(sheet, spreadsheet, shippingAmount)
  }
}

/**
 * This function ...
 * 
 * @author Jarren Ralf
 */
function packingSlip_HFPU()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const activeSheet = spreadsheet.getActiveSheet();

  if (activeSheet.getSheetName() !== 'Invoice' && activeSheet.getSheetName() !== 'Packing Slip')
  {
    spreadsheet.toast('You must be on the Invoice sheet to run this function. Please try again.')
    spreadsheet.getSheetByName('Invoice').activate();
  }
  else
  {
    const sheet = spreadsheet.getSheetByName('Invoice')
    const currentOrder = sheet.getSheetValues(1, 9, 1, 1)[0][0]
    const statusPage = spreadsheet.getSheetByName('Status Page')
    const range = statusPage.getRange(3, 1, statusPage.getLastRow() - 2, 4);
    var values = range.getValues();

    const shippingAmount = sheet.getSheetValues(4, 9, 1, 1)[0][0];
    const activeOrdersPage = spreadsheet.getSheetByName('All_Active_Orders')
    const activeOrdersValues = activeOrdersPage.getSheetValues(1, 1, activeOrdersPage.getLastRow(), activeOrdersPage.getLastColumn())
    const values_ExportPage = [activeOrdersValues[0]]; // The shopify data used to create the export data for Adagio; initialize with the header

    for (var j = 1; j < activeOrdersValues.length; j++)
    {
      if (activeOrdersValues[j][0] === currentOrder)
      {
        values_ExportPage.push(activeOrdersValues[j])
        break;
      }
    }

    for (var i = 0; i < values.length; i++)
    {
      if (values[i][0] === currentOrder)
      {
        values[i][3] = 'Staged for Pickup - '

        switch (activeOrdersValues[j][14])
        {
          case 'Richmond':
            values[i][3] += activeOrdersValues[j][14]
            break;
          case 'Parksville':
            values[i][3] += activeOrdersValues[j][14]
            break;
          case 'Prince Rupert':
            values[i][3] += 'Rupert'
            break;
        }

        break;
      }
    }

    range.setValues(values)

    var lastCol = 9;
    var numPages = sheet.getSheetValues(50, lastCol, 1, 1)[0][0];

    if (isBlank(numPages))
      var itemValues_PackingSlip = sheet.getSheetValues(17, 1, 32, lastCol - 1);
    else
    {
      numPages = numPages.split(' of ')[1]
      var itemValues_PackingSlip = sheet.getSheetValues(17, 1, 32, lastCol - 1);

      for (var p = 0; p < numPages; p++)
        itemValues_PackingSlip.push(...sheet.getSheetValues(59 + p*49, 1, 39, lastCol - 1))
    }

    const completedOrdersPage = spreadsheet.getSheetByName('Completed Orders');
    const lastRow = completedOrdersPage.getLastRow();
  
    const ordersOnCompletePage = (lastRow === 0) ? [[currentOrder + ' - Hold For Pick Up']] : 
      Array.from(new Set(completedOrdersPage.getSheetValues(1, 1, completedOrdersPage.getLastRow(), 1)
        .concat([[currentOrder + ' - Hold For Pick Up']]).sort((a, b) => (a[0] < b[0]) ? 1 : -1).map(JSON.stringify)), JSON.parse)

    completedOrdersPage.getRange(1, 1, ordersOnCompletePage.length).setValues(ordersOnCompletePage)

    exportData(values_ExportPage, spreadsheet.getSheetByName('Export'), spreadsheet, shippingAmount, itemValues_PackingSlip);
    statusPage.activate()
    applyFormattingToPackingSlip(sheet, spreadsheet, shippingAmount)
  }
}

/**
 * This function ...
 * 
 * @param {Event} e : The event object.
 * @throws Throws an error if the script doesn't run
 * @author Jarren Ralf
 */
function processImportedData(e)
{
  try
  {
    var spreadsheet = e.source;
    var sheets = spreadsheet.getSheets();
    var info, numRows = 0, numCols = 1, maxRow = 2, maxCol = 3;

    for (var sheet = 0; sheet < sheets.length; sheet++) // Loop through all of the sheets in this spreadsheet and find the new one
    {
      info = [
        sheets[sheet].getLastRow(),
        sheets[sheet].getLastColumn(),
        sheets[sheet].getMaxRows(),
        sheets[sheet].getMaxColumns()
      ]

      // A new sheet is imported by File -> Import -> Insert new sheet(s) - The left disjunct is for a csv and the right disjunct is for an excel file
      if ((info[maxRow] - info[numRows] === 2 && info[maxCol] - info[numCols] === 2) || (info[maxRow] === 1000 && info[maxCol] === 26 && info[numRows] !== 0 && info[numCols] !== 0)) 
      {
        const values = sheets[sheet].getSheetValues(1, 1, info[numRows], info[numCols]); // This is the shopify order data
        const numOrders = updateStatusPage(values, spreadsheet);
        updateActiveOrderPage(values, spreadsheet);

        if (numOrders === 1)
          updatePackingSlip(values, info[numRows], info[numCols], spreadsheet)

        if (sheets[sheet].getSheetName().substring(0, 7) !== "Copy Of") // Don't delete the sheets that are duplicates
          spreadsheet.deleteSheet(sheets[sheet]) // Delete the new sheet that was created

        break;
      }
    }
  }
  catch (err)
  {
    spreadsheet.deleteSheet(sheets[sheet]) // Delete the new sheet(s) that was created

    var error = err['stack'];
    Logger.log(error);
    Browser.msgBox('Please contact the spreadsheet owner and let them know what action you were performing that lead to the following error: ' + error)
    throw new Error(error);
  }
}

/**
 * This function ...
 * 
 * @param {Event} e : The event object.
 * @throws Throws an error if the script doesn't run
 * @author Jarren Ralf
 */
function reformatPackingSlip(e)
{
  try
  {
    var spreadsheet = e.source;
    var sheet = spreadsheet.getActiveSheet();
    var sheetName = sheet.getSheetName();

    if (sheetName === 'Invoice')
    {
      const numRowsPerPage = 49;
      const numItemsOnPageOne = 32;
      const lastRow = sheet.getLastRow()
      const lastCol = 9;
      const values = sheet.getDataRange().getValues();
      const items = values.filter(val => val[0].toString().substr(-1) === 'x')

      if (lastRow <= numRowsPerPage) // Only 1 page
      {
        sheet.insertRowsAfter(numRowsPerPage - numItemsOnPageOne + items.length - 1, numRowsPerPage - lastRow + 1)
          .getRange(numRowsPerPage - numItemsOnPageOne + items.length, 3, numItemsOnPageOne - items.length, 5).mergeAcross()
          .offset(0, -2, numItemsOnPageOne - items.length, lastCol)
          .setBorder(false, true, null, true, true, false)
      }
      else // Multipage
      {
        const firstRow_PageOne = 16
        const firstRow_PageTwo = 58
        const numItemsPerPage = 39;
        const numPages = sheet.getRange(lastRow, lastCol).getValue().split(' of ')[1];
        const numDeletedRows = numPages*numRowsPerPage + 1 - lastRow;
        const numPagesRequired = Math.ceil((items - numItemsOnPageOne) / numItemsPerPage) + 1

        if (numPagesRequired < numPages) // We must reduce the number of pages
        {

        }
        else // The number of pages stay the same
        {
          for (var row = 0; row < numItemsOnPageOne; row++)
          {
            if (isBlank(values[firstRow_PageOne + row][0])) 
            {
              sheet.insertRowsBefore(firstRow_PageOne + row, numItemsOnPageOne - row)
                .getRange(firstRow_PageOne + row, 3, numItemsOnPageOne - row, 5).mergeAcross()
                .offset(0, -2, numItemsOnPageOne - row + 1, lastCol)
                .setValues(items.slice(row - 1, numItemsOnPageOne))
              break;
            }
          }

          if (numDeletedRows > numItemsOnPageOne - row)
          {
            // for (var page = 0; page < numPages - 1; page++)
            // {
            //   for (var row = firstRow_PageTwo + numRowsPerPage*page; row < numItemsPerPage; row++)
            //   {
            //     if (isBlank(values[firstRow + row][0]))
            //     {
            //       sheet.insertRowsBefore(firstRow + row, numItemsOnPageOne - row)
            //         .getRange(firstRow + row, 3, numItemsOnPageOne - row, 5)
            //         .mergeAcross()
            //         .offset(0, -2, numItemsOnPageOne - row + 1, lastCol)
            //         .setValues(items.slice(row - 1, numItemsOnPageOne))
            //       break;
            //     }
            //   }
            // }
          }
          else
            Logger.log('Restored all of the rows!')
        }
      }
    }
  }
  catch (err)
  {
    var error = err['stack'];
    Logger.log(error);
    Browser.msgBox('Please contact the spreadsheet owner and let them know what action you were performing that lead to the following error: ' + error)
    throw new Error(error);
  }
}

/**
 * This function finds the orders that have a COMPLETED status and removes them.
 * 
 * @author Jarren Ralf
 */
function removeCompleteOrdersButton()
{
  var isComplete, isCancelled, isPickedUp, completedOrders = [], pickedUpOrders = [], cancelledOrders = [];
  const spreadsheet = SpreadsheetApp.getActive()
  const sheet = SpreadsheetApp.getActiveSheet();
  const numCols = sheet.getLastColumn();
  const values = sheet.getSheetValues(3, 1, sheet.getLastRow() - 2, numCols);

  const remainingOrders = values.filter(v => {
    isComplete  = v[3] == 'COMPLETED';
    isPickedUp  = v[3] == 'PICKED UP IN STORE';
    isCancelled = v[3] == 'CANCELLED';
    
    if (isComplete)
      completedOrders.push(v[0]); // Compile a list of completed order numbers
    
    if (isPickedUp)
      pickedUpOrders.push(v[0]); // Compile a list of picked up order numbers

    if (isCancelled)
      cancelledOrders.push(v[0]); // Compile a list of cancelled order numbers

    return !(isComplete || isPickedUp || isCancelled); // Not complete and not cancelled and not picked up, and therefore the remaining orders
  });

  const numOrders = remainingOrders.length;
  const numberFormats = new Array(numOrders).fill(['@', '#', ...new Array(numCols - 3).fill('@'), "dd MMM yyyy"])

  sheet.getRange(3, 1, numOrders, numCols).setNumberFormats(numberFormats).setValues(remainingOrders)
  sheet.deleteRows(numOrders + 3, values.length - numOrders) // Delete the last rows

  var isCompletedOrder, isPickedUpOrder, isCancelledOrder;
  const activeOrdersPage = spreadsheet.getSheetByName('All_Active_Orders')
  const numCols_OrdersPage = activeOrdersPage.getLastColumn();
  const values_OrdersPage = activeOrdersPage.getSheetValues(2, 1, activeOrdersPage.getLastRow() - 1, numCols_OrdersPage);
  const values_ExportPage = [activeOrdersPage.getSheetValues(1, 1, 1, numCols_OrdersPage)[0]]; // The shopify data used to create the export data for Adagio; initialize with the header

  const remainingOrders_OrdersPage = values_OrdersPage.filter(v => {
    isCompletedOrder = completedOrders.includes(v[0]);
    isPickedUpOrder  = pickedUpOrders.includes(v[0]);
    isCancelledOrder = cancelledOrders.includes(v[0]);

    if (isCompletedOrder)
      values_ExportPage.push(v);

    return !(isCompletedOrder || isPickedUpOrder || isCancelledOrder);
  });
    
  const numRows = remainingOrders_OrdersPage.length;

  activeOrdersPage.getRange(2, 1, numRows, numCols_OrdersPage).setValues(remainingOrders_OrdersPage)

  const completedOrdersPage = spreadsheet.getSheetByName('Completed Orders');
  const lastRow = completedOrdersPage.getLastRow();
  
  const ordersOnCompletePage = (lastRow === 0) ? completedOrders.map(v => [v + ' - Completed'])
      .concat(pickedUpOrders.map(v => [v + ' - Completed']))
      .concat(cancelledOrders.map(v => [v + ' - Cancelled']))
      .sort((a, b) => (a[0] < b[0]) ? -1 : 1) : 
    completedOrdersPage.getSheetValues(1, 1, completedOrdersPage.getLastRow(), 1)
      .concat(completedOrders.map(v => [v + ' - Completed']), pickedUpOrders.map(v => [v + ' - Completed']), cancelledOrders.map(v => [v + ' - Cancelled']))
      .sort((a, b) => (a[0] < b[0]) ? 1 : -1);

  completedOrdersPage.getRange(1, 1, ordersOnCompletePage.length).setValues(ordersOnCompletePage)

  if (values_OrdersPage.length !== numRows)
  {
    activeOrdersPage.deleteRows(numRows + 2, values_OrdersPage.length - numRows) // Delete the last rows
    exportData(values_ExportPage, spreadsheet.getSheetByName('Export'), spreadsheet);
  }
}

/**
 * This function resets the array formulas on the Packing Slip that sets the items and quantities derived from the Invoice sheet.
 * 
 * @author Jarren Ralf
 */
function resetArrayFormulaPackingSlip()
{
  SpreadsheetApp.getActive().getSheetByName('Packing Slip').getRange(15, 1, 1, 8).setFormulas([[
    '=ARRAYFORMULA(Invoice!$C17:C$34&char(10)&if(Invoice!$C17:C$34=\"\",\"\",\"Sku# \"&Invoice!$B17:B$34))', '', '', '', '', '', '',
    '=ARRAYFORMULA(IFERROR(query(SPLIT(Invoice!A17:$A34, \" \"), \"SELECT Col1\"),\"\"))'
  ]])
}

/**
 * This function...
 * 
 * @author Jarren Ralf
 */
function showItems(fulfilmentStatus)
{
  const spreadsheet = SpreadsheetApp.getActive()
  const packingSlip = spreadsheet.getSheetByName('Invoice').activate();
  const shopifyData = spreadsheet.getSheetByName('Last_Import').getDataRange().getValues();
  const col = 9; // Number of columns on the packing slip
  const orderNumber = packingSlip.getSheetValues(1, col, 1, 1)[0][0];

  if (orderNumber !== shopifyData[1][0])
    Browser.msgBox('The order number on this Invoice does not match the Last_Import page.\n\nPlease use File -> Import to upload the desired Invoice.')
  else
  {
    const numRowsPerPage = 49;
    const numItemsOnPageOne = 32;
    const numItemsPerPage = 39; // Starting with page 2
    const numPages = Math.ceil((shopifyData.length - numItemsOnPageOne - 1) / numItemsPerPage) + 1

    switch (fulfilmentStatus)
    {
      case 'all':
        var itemInformation = shopifyData.map(val => [val[16].toString() + ' x', val[20], val[17], null, null, null, null, val[18] - val[59], (val[18] - val[59])*val[16]]); // All of the items
        itemInformation.shift()
        break;
      case 'pending':
      case 'fulfilled':
      case 'unfulfilled':
        var itemInformation = shopifyData.filter(val => val[23] === fulfilmentStatus)
          .map(val => [val[16].toString() + ' x', val[20], val[17], null, null, null, null, val[18] - val[59], (val[18] - val[59])*val[16]])
        break;
      case 'pending & fulfilled':
        var itemInformation = shopifyData.filter(val => val[23] === 'pending' || val[23] === 'fulfilled')
          .map(val => [val[16].toString() + ' x', val[20], val[17], null, null, null, null, val[18] - val[59], (val[18] - val[59])*val[16]])
        break;
    }
    if (itemInformation.length !== 0)
    {
      if (numPages >= 2) // Two or more pages 
      {
        packingSlip.getRange(numRowsPerPage + 1, col).setHorizontalAlignment('right').setValue('Page 1 of ' + numPages) // Put the page number on the bottom of page one
        packingSlip.insertRowsAfter(numRowsPerPage + 1, (numPages - 1)*(numRowsPerPage))

        if (packingSlip.getMaxRows() > numPages*numRowsPerPage + 1)
          packingSlip.deleteRows(numPages*numRowsPerPage + 2, packingSlip.getMaxRows() - numPages*numRowsPerPage - 1)

        packingSlip.getRange(16, 1, numItemsOnPageOne + 1, col) // Item information and formatting
          .setBorder(true, true, true, true, true, false).setFontColor('black').setFontFamily('Arial')
          .setHorizontalAlignments(new Array(numItemsOnPageOne + 1).fill(['center', 'center', 'left', 'left', 'left', 'left', 'left', 'right', 'right']))
          .setVerticalAlignments(new Array(numItemsOnPageOne + 1).fill(new Array(col).fill('middle')))
          .setFontSizes([new Array(col).fill(12), ...new Array(numItemsOnPageOne).fill(new Array(col).fill(9))])
          .setFontWeights([new Array(col).fill('bold'), ...new Array(numItemsOnPageOne).fill(new Array(col).fill('normal'))])
          .setNumberFormats([new Array(col).fill('@'), ...new Array(numItemsOnPageOne).fill(['@', '@', '@', '@', '@', '@', '@', '$#,##0.00', '$#,##0.00'])])
          .setValues([['Qty', 'SKU', 'Item', null, null, null, null, 'Price', 'Total'], ...itemInformation.slice(0, numItemsOnPageOne)])
        
        const pntAddress = packingSlip.getRange(4, 1).getRichTextValue();
        const emailHyperLink = packingSlip.getRange(numRowsPerPage + 1, 1).getRichTextValue();
        const invoiceHeaderValues = packingSlip.getRange(1, col - 1, 6).getValues().map((v, i) => [v[0], '=I' + (i + 1)])
        var subtotalAmount = '=SUM(Item_Totals_Page_1', rangeName = '';

        for (var n = 0; n < numPages - 1; n++)
        {
          var N = numRowsPerPage*n;

          rangeName = 'Item_Totals_Page_' + (n + 2);
          spreadsheet.setNamedRange(rangeName, packingSlip.getRange(numRowsPerPage + 10 + N, col, numItemsPerPage))
          subtotalAmount += ',' + rangeName

          packingSlip.setRowHeight(numRowsPerPage + 3 + N, 40)
            .setRowHeight(numRowsPerPage + 8 + N, 10)
            .setRowHeight(numRowsPerPage + 49 + N, 10)
            .getRange(numRowsPerPage + 9 + N, 3, numItemsPerPage + 1, 5).mergeAcross(); // Item (Description Field)
          packingSlip.getRange(numRowsPerPage + 2 + N, 1, 3, 3).merge().setFormula('=A1'); // PNT Logo in Header
          packingSlip.getRange(numRowsPerPage + 5 + N, 1, 3, 3).merge().setVerticalAlignment('middle').setHorizontalAlignment('left').setRichTextValue(pntAddress); // PNT Address in header
          packingSlip.getRange(numRowsPerPage + 50 + N, 1, 1, 5).merge().setRichTextValue(emailHyperLink) // Email Hyperlink at bottom of each page
          packingSlip.getRange(numRowsPerPage + 50 + N, col).setHorizontalAlignment('right').setValue('Page ' + (n + 2) + ' of ' + numPages) // Page number for the bottom of each page

          packingSlip.getRange(numRowsPerPage + 2 + N, 8, 6, 2) // Invoice header data
            .setFontColor('black').setFontFamily('Arial')
            .setFontSizes([[10, 10],[10, 9], ...new Array(4).fill([10, 10])])
            .setFontWeights(new Array(6).fill(['bold', 'normal']))
            .setHorizontalAlignments([['right', 'center'], ['right', 'center'], ...new Array(4).fill(['right', 'right'])])
            .setVerticalAlignments([['middle', 'middle'], ['top', 'top'], ...new Array(4).fill(['middle', 'middle'])])
            .setNumberFormats([['@', '@'], ['@', 'dd MMM yyyy'], ...new Array(4).fill(['@', '$#,##0.00'])])
            .setValues(invoiceHeaderValues) // Header Values

          if (n != numPages - 2)
            packingSlip.getRange(numRowsPerPage + 9 + N, 1, numItemsPerPage + 1, col) // Item information and formatting
              .setBorder(true, true, true, true, true, false).setFontColor('black').setFontFamily('Arial')
              .setHorizontalAlignments(new Array(numItemsPerPage + 1).fill(['center', 'center', 'left', 'left', 'left', 'left', 'left', 'right', 'right']))
              .setVerticalAlignments(new Array(numItemsPerPage + 1).fill(new Array(col).fill('middle')))
              .setFontSizes([new Array(col).fill(12), ...new Array(numItemsPerPage).fill(new Array(col).fill(9))])
              .setFontWeights([new Array(col).fill('bold'), ...new Array(numItemsPerPage).fill(new Array(col).fill('normal'))])
              .setNumberFormats([new Array(col).fill('@'), ...new Array(numItemsPerPage).fill(['@', '@', '@', '@', '@', '@', '@', '$#,##0.00', '$#,##0.00'])])
              .setValues([['Qty', 'SKU', 'Item', null, null, null, null, 'Price', 'Total'], 
                ...itemInformation.slice(numItemsOnPageOne + numItemsPerPage*n, numItemsOnPageOne + numItemsPerPage*(n + 1))])
          else // Last Page
            packingSlip.getRange(numRowsPerPage + 9 + N, 1, numItemsPerPage + 1, col) // Item information and formatting
              .setBorder(true, true, true, true, true, false).setFontColor('black').setFontFamily('Arial')
              .setHorizontalAlignments(new Array(numItemsPerPage + 1).fill(['center', 'center', 'left', 'left', 'left', 'left', 'left', 'right', 'right']))
              .setVerticalAlignments(new Array(numItemsPerPage + 1).fill(new Array(col).fill('middle')))
              .setFontSizes([new Array(col).fill(12), ...new Array(numItemsPerPage).fill(new Array(col).fill(9))])
              .setFontWeights([new Array(col).fill('bold'), ...new Array(numItemsPerPage).fill(new Array(col).fill('normal'))])
              .setNumberFormats([new Array(col).fill('@'), ...new Array(numItemsPerPage).fill(['@', '@', '@', '@', '@', '@', '@', '$#,##0.00', '$#,##0.00'])])
              .setValues([['Qty', 'SKU', 'Item', null, null, null, null, 'Price', 'Total'], 
                ...itemInformation.slice(numItemsOnPageOne + numItemsPerPage*n, itemInformation.length), 
                ...new Array(numItemsOnPageOne + numItemsPerPage*(n + 1) - itemInformation.length).fill(new Array(col).fill(''))])
        }

        // setValues of total
        subtotalAmount += ')';
        packingSlip.getRange(3, col).setFormula(subtotalAmount)
      }
      else
      {
        packingSlip.getRange(16, 1, numItemsOnPageOne + 1, col) // Item information and formatting
          .setBorder(true, true, true, true, true, false).setFontColor('black').setFontFamily('Arial')
          .setHorizontalAlignments(new Array(numItemsOnPageOne + 1).fill(['center', 'center', 'left', 'left', 'left', 'left', 'left', 'right', 'right']))
          .setVerticalAlignments(new Array(numItemsOnPageOne + 1).fill(new Array(col).fill('middle')))
          .setFontSizes([new Array(col).fill(12), ...new Array(numItemsOnPageOne).fill(new Array(col).fill(9))])
          .setFontWeights([new Array(col).fill('bold'), ...new Array(numItemsOnPageOne).fill(new Array(col).fill('normal'))])
          .setNumberFormats([new Array(col).fill('@'), ...new Array(numItemsOnPageOne).fill(['@', '@', '@', '@', '@', '@', '@', '$#,##0.00', '$#,##0.00'])])  
          .setValues([['Qty', 'SKU', 'Item', null, null, null, null, 'Price', 'Total'],
            ...itemInformation.slice(0, itemInformation.length), 
            ...new Array(numItemsOnPageOne - itemInformation.length).fill(new Array(col).fill(''))])

        if (packingSlip.getMaxRows() > numRowsPerPage + 1)
          packingSlip.deleteRows(numRowsPerPage + 2, packingSlip.getMaxRows() - numRowsPerPage - 1) // One page, delete the extra rows
          
        packingSlip.getRange(numRowsPerPage + 1, col).setValue('') // Set the page number blank

        var namedRange = '';

        packingSlip.getNamedRanges().map(rng => {
          namedRange = rng.getName()

          if (namedRange[namedRange.length - 1] !== '1' && !isNaN(parseInt(namedRange[namedRange.length - 1]))) // If the range ends with a number that is not 1, then remove it
            rng.remove();
        })

        packingSlip.getRange(3, col).setFormula('=SUM(Item_Totals_Page_1)')
      }
    }
    else
    {
      if (packingSlip.getMaxRows() > numRowsPerPage + 1)
        packingSlip.deleteRows(numRowsPerPage + 2, packingSlip.getMaxRows() - numRowsPerPage - 1) // One page, delete the extra rows

      packingSlip.getRange(17, 1, numItemsOnPageOne, col).setValue('') // Set all of item information to blank
      packingSlip.getRange(50, col).setValue('') // Only 1 page so set the page number to blank
    }
    
    spreadsheet.getRangeByName('ShippingAmount').activate();
  }
}

/**
 * This function...
 * 
 * @author Jarren Ralf
 */
function showItems_all()
{
  showItems('all')
}

/**
 * This function...
 * 
 * @author Jarren Ralf
 */
function showItems_fulfilled()
{
  showItems('fulfilled')
}

/**
 * This function...
 * 
 * @author Jarren Ralf
 */
function showItems_pending()
{
  showItems('pending')
}

/**
 * This function...
 * 
 * @author Jarren Ralf
 */
function showItems_pending_AND_fulfilled()
{
  showItems('pending & fulfilled')
}

/**
 * This function...
 * 
 * @author Jarren Ralf
 */
function showItems_unfulfilled()
{
  showItems('unfulfilled')
}

/**
 * This function takes the given string and makes sure that each word in the string has a capitalized 
 * first letter followed by lower case.
 * 
 * @param {String} str : The given string
 * @return {String} The output string with proper case
 * @author Jarren Ralf
 */
function toProper(str)
{
  return capitalizeSubstrings(capitalizeSubstrings(str, '-'), ' ');
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
 * This function updates ...
 * 
 * @param {Object[][]}  shopifyData : The data downloaded from shopify.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @return {Number} Returns the number of orders in the particular shopify data set.
 * @author Jarren Ralf
 */
function updateActiveOrderPage(shopifyData, spreadsheet)
{
  const header = shopifyData.shift() // Remove the header
  const listOfOrderNumbers_ContainsDuplicates = shopifyData.map(val => val[0])
  const listOfImportedOrders =  [...new Set(listOfOrderNumbers_ContainsDuplicates)] 
  const activeOrdersPage = spreadsheet.getSheetByName('All_Active_Orders');
  const activeOrdersData = activeOrdersPage.getSheetValues(2, 1, activeOrdersPage.getLastRow() - 1, activeOrdersPage.getLastColumn()).filter(val => !listOfImportedOrders.includes(val[0]));
  const updatedData = activeOrdersData.concat(shopifyData)
  shopifyData.unshift(header)
  activeOrdersPage.getRange(2, 1, updatedData.length, activeOrdersData[0].length).setNumberFormat('@').setValues(updatedData)
}

/**
 * This function updates the 
 * 
 * @param {Object[][]}  shopifyData : The data downloaded from shopify.
 * @param {Object[][]}    numRows   : The number of rows for the shopify data.
 * @param {Object[][]}    numCols   : The number of columns for the shopify data.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @author Jarren Ralf
 */
function updatePackingSlip(shopifyData, numRows, numCols, spreadsheet)
{
  const packingSlip = spreadsheet.getSheetByName('Invoice');
  const col = 9; // Number of columns on the packing slip
  const numRowsPerPage = 49;
  const numItemsOnPageOne = 32;
  const numItemsPerPage = 39; // Starting with page 2
  const header = shopifyData.shift()
  const numPages = Math.ceil((shopifyData.length - numItemsOnPageOne) / numItemsPerPage) + 1
  
  spreadsheet.getRangeByName('Hidden_Checkbox').uncheck(); // This is the checkbox on the Packing Slip that adds 10%
  const calculator = spreadsheet.getSheetByName('Calculator');
  calculator.getRange(2, 3, 2).setFormulas([['=SubtotalAmount'], ['=ShippingAmount']])
  calculator.getRange('D2:D10').uncheck()
  const checkboxRange = spreadsheet.getRangeByName('Checkboxes'); // These are the checkboxes that control the taxation rate
  const checks = checkboxRange.getValues()

  shopifyData[0][ 1] = shopifyData[0][1].toString().toLowerCase(); // Email
  shopifyData[0][24] = toProper(shopifyData[0][24]); // Billing Name
  shopifyData[0][34] = toProper(shopifyData[0][34]); // Shipping Name
  shopifyData[0][28] = toProper(shopifyData[0][28]); // Billing Company Name
  shopifyData[0][38] = toProper(shopifyData[0][38]); // Shipping Company Name Name
  shopifyData[0][26] = toProper(shopifyData[0][26]); // Billing Address
  shopifyData[0][36] = toProper(shopifyData[0][36]); // Shipping Address
  shopifyData[0][29] = toProper(shopifyData[0][29]); // Billing City
  shopifyData[0][39] = toProper(shopifyData[0][39]); // Shipping City
  shopifyData[0][31] = shopifyData[0][31].toString().toUpperCase(); // Billing Province
  shopifyData[0][41] = shopifyData[0][41].toString().toUpperCase(); // Shipping Province
  shopifyData[0][30] = shopifyData[0][30].toString().toUpperCase(); // Billing Postal
  shopifyData[0][40] = shopifyData[0][40].toString().toUpperCase(); // Shipping Postal
  shopifyData[0][30] = formatPostalCode(shopifyData[0][30]); // Billing Postal
  shopifyData[0][40] = formatPostalCode(shopifyData[0][40]); // Shipping Postal
  shopifyData[0][33] = formatPhoneNumber(shopifyData[0][33]); // Billing Phone Number
  shopifyData[0][43] = formatPhoneNumber(shopifyData[0][43]); // Shipping Phone Number

  // Check the shipping country and province, then set the taxes accordingly by checking the appropriate box
  if (isBlank(shopifyData[0][41])) // Blank means the item is a pick up in BC, therefore charge 12%
  {
    checks[0][0] = 0.12;
    spreadsheet.getRangeByName('ShippingAmount').setValue(0);
  }
  else
  {
    if (shopifyData[0][42] !== 'CA')
      checks[5][0] = 0;
    else
    {
      if (shopifyData[0][41] === 'BC') 
        checks[0][0] = 0.12;
      else if (shopifyData[0][41] === 'AB' || shopifyData[0][41] === 'NT' || shopifyData[0][41] === 'NU' || 
               shopifyData[0][41] === 'YT' || shopifyData[0][41] === 'QC' || shopifyData[0][41] === 'MB' ||
               shopifyData[0][41] === 'SK')
        checks[1][0] = 0.05;
      else if (shopifyData[0][41] === 'NS' || shopifyData[0][41] === 'NB' || shopifyData[0][41] === 'NL' || shopifyData[0][41] === 'PE')
        checks[2][0] = 0.15;
      else if (shopifyData[0][41] === 'ON')
        checks[3][0] = 0.13;
      // else if (shopifyData[0][41] === 'SK')
      //   checks[4][0] = 0.11;
    }
  }

  const tax = 1 + checks.reduce((acc, val) => acc + val[0] , 0);
  const updatedSubTotal = shopifyData.reduce((acc, val) => acc + val[16]*val[18], 0);
  const freightCost = shopifyData[0][11] - shopifyData[0][51] - updatedSubTotal*tax; // Total - outstanding balance - subtotal*tax
  const shippingMethod = packingSlip.getRange(14, 2);
  const shippingCost = packingSlip.getRange(4, col);
  const formattedDate = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), "dd MMMM yyyy");
  const shipDate = 'Ship Date: ' + formattedDate;
  const boldTextStyle = SpreadsheetApp.newTextStyle().setBold(true).setFontSize(12).build();
  const normalTextStyle = SpreadsheetApp.newTextStyle().setBold(false).setFontSize(10).build();
  const shipDate_RichText = SpreadsheetApp.newRichTextValue().setText('Ship Date: ' + formattedDate).setTextStyle(0, 10, boldTextStyle).setTextStyle(10, shipDate.length, normalTextStyle).build();

  packingSlip.getRange(14, 4).setValue('')
  packingSlip.getRange(14, 7).setRichTextValue(shipDate_RichText)
  shippingCost.setValue(freightCost)

  // Check the shipping method and make the relevant changes
  switch (shopifyData[0][14])
  {
    case 'Richmond':
      shopifyData[0][34] = shopifyData[0][24]; // Name
      shopifyData[0][38] = 'Pacific Net & Twine';
      shopifyData[0][36] = '3731 Moncton Street';
      shopifyData[0][39] = shopifyData[0][14]; // City
      shopifyData[0][41] = 'BC';
      shopifyData[0][40] = 'V7E 3A5';
      shopifyData[0][43] = shopifyData[0][33]; // Phone Number;
      shippingMethod.setValue('Pick Up')
      calculator.getRange('D9').check() // Check the box on calculator for pick up
      shippingCost.setValue(0)
      break;
    case 'Parksville':
      shopifyData[0][34] = shopifyData[0][24]; // Name
      shopifyData[0][38] = 'Pacific Net & Twine';
      shopifyData[0][36] = '1380 Alberni Hwy';
      shopifyData[0][39] = shopifyData[0][14]; // City
      shopifyData[0][41] = 'BC';
      shopifyData[0][40] = 'V9P 2C9';
      shopifyData[0][43] = shopifyData[0][33]; // Phone Number;
      shippingMethod.setValue('Pick Up')
      calculator.getRange('D9').check() // Check the box on calculator for pick up
      shippingCost.setValue(0)
      break;
    case 'Prince Rupert':
      shopifyData[0][34] = shopifyData[0][24]; // Name
      shopifyData[0][38] = 'Pacific Net & Twine';
      shopifyData[0][36] = '125 1st Ave W';
      shopifyData[0][39] = shopifyData[0][14]; // City
      shopifyData[0][41] = 'BC';
      shopifyData[0][40] = 'V8J 4K8';
      shopifyData[0][43] = shopifyData[0][33]; // Phone Number;
      shippingMethod.setValue('Pick Up')
      calculator.getRange('D9').check() // Check the box on calculator for pick up
      shippingCost.setValue(0)
      break;
    case 'Post Lettermail':
    case 'CDA Post Lettermail':
      calculator.getRange('D10').check() // Check the box on calculator for lettermail
      shippingMethod.setValue('Post Lettermail')
      break;
    case 'Post Expedited Parcel':
    case 'CDA Post Expedited Parcel':
      shippingMethod.setValue('Post Expedited Parcel')
      break;
    case 'Post Tracked Packet':
    case 'CDA Post Tracked Packet':
      shippingMethod.setValue('Post Tracked Packet')
      break;
    case 'CDA Post Xpress Post':
    case 'Post Xpress Post':
      shippingMethod.setValue('Post Xpress Post')
      break;
    case 'Purolator Ground':
    case 'UPS Standard':
    case 'UPS Express':
    case 'Lower Mainland Fast Freight':
      shippingMethod.setValue(shopifyData[0][14])
      break;
    default:
      shippingMethod.setValue('Select Shipping Method')
  }

  const itemInformation = shopifyData.map(val => [val[16].toString() + ' x', val[20], val[17], null, null, null, null, val[18] - val[59]/val[16], val[18]*val[16] - val[59]]) // Subtract Lineitem discount
  shopifyData.unshift(header);
  spreadsheet.getSheetByName('Last_Import').clearContents().getRange(1, 1, numRows, numCols).setValues(shopifyData) // Put all of the imported data on the Last_Import sheet

  if (numPages >= 2) // Two or more pages 
  {
    packingSlip.getRange(numRowsPerPage + 1, col).setHorizontalAlignment('right').setValue('Page 1 of ' + numPages) // Put the page number on the bottom of page one
    packingSlip.insertRowsAfter(numRowsPerPage + 1, (numPages - 1)*(numRowsPerPage))

    if (packingSlip.getMaxRows() > numPages*numRowsPerPage + 1)
      packingSlip.deleteRows(numPages*numRowsPerPage + 2, packingSlip.getMaxRows() - numPages*numRowsPerPage - 1)

    packingSlip.getRange(16, 1, numItemsOnPageOne + 1, col) // Item information and formatting
      .setBorder(true, true, true, true, true, false).setFontColor('black').setFontFamily('Arial')
      .setHorizontalAlignments(new Array(numItemsOnPageOne + 1).fill(['center', 'center', 'left', 'left', 'left', 'left', 'left', 'right', 'right']))
      .setVerticalAlignments(new Array(numItemsOnPageOne + 1).fill(new Array(col).fill('middle')))
      .setFontSizes([new Array(col).fill(12), ...new Array(numItemsOnPageOne).fill(new Array(col).fill(9))])
      .setFontWeights([new Array(col).fill('bold'), ...new Array(numItemsOnPageOne).fill(new Array(col).fill('normal'))])
      .setNumberFormats([new Array(col).fill('@'), ...new Array(numItemsOnPageOne).fill(['@', '@', '@', '@', '@', '@', '@', '$#,##0.00', '$#,##0.00'])])
      .setValues([['Qty', 'SKU', 'Item', null, null, null, null, 'Price', 'Total'], ...itemInformation.slice(0, numItemsOnPageOne)])
    
    const pntAddress = packingSlip.getRange(4, 1).getRichTextValue();
    const emailHyperLink = packingSlip.getRange(numRowsPerPage + 1, 1).getRichTextValue()
    const invoiceHeaderValues = packingSlip.getRange(1, col - 1, 6).getValues().map((v, i) => [v[0], '=I' + (i + 1)])
    var subtotalAmount = '=SUM(Item_Totals_Page_1', rangeName = '';

    for (var n = 0; n < numPages - 1; n++)
    {
      var N = numRowsPerPage*n;

      rangeName = 'Item_Totals_Page_' + (n + 2);
      spreadsheet.setNamedRange(rangeName, packingSlip.getRange(numRowsPerPage + 10 + N, col, numItemsPerPage))
      subtotalAmount += ',' + rangeName

      packingSlip.setRowHeight(numRowsPerPage + 3 + N, 40)
        .setRowHeight(numRowsPerPage + 8 + N, 10)
        .setRowHeight(numRowsPerPage + 49 + N, 10)
        .getRange(numRowsPerPage + 9 + N, 3, numItemsPerPage + 1, 5).mergeAcross(); // Item (Description Field)
      packingSlip.getRange(numRowsPerPage + 2 + N, 1, 3, 3).merge().setFormula('=A1'); // PNT Logo in Header
      packingSlip.getRange(numRowsPerPage + 5 + N, 1, 3, 3).merge().setVerticalAlignment('middle').setHorizontalAlignment('left').setRichTextValue(pntAddress); // PNT Address in header
      packingSlip.getRange(numRowsPerPage + 50 + N, 1, 1, 5).merge().setRichTextValue(emailHyperLink) // Email Hyperlink at bottom of each page
      packingSlip.getRange(numRowsPerPage + 50 + N, col).setHorizontalAlignment('right').setValue('Page ' + (n + 2) + ' of ' + numPages) // Page number for the bottom of each page

      packingSlip.getRange(numRowsPerPage + 2 + N, 8, 6, 2) // Invoice header data
        .setFontColor('black').setFontFamily('Arial')
        .setFontSizes([[10, 10],[10, 9], ...new Array(4).fill([10, 10])])
        .setFontWeights(new Array(6).fill(['bold', 'normal']))
        .setHorizontalAlignments([['right', 'center'], ['right', 'center'], ...new Array(4).fill(['right', 'right'])])
        .setVerticalAlignments([['middle', 'middle'], ['top', 'top'], ...new Array(4).fill(['middle', 'middle'])])
        .setNumberFormats([['@', '@'], ['@', 'dd MMM yyyy'], ...new Array(4).fill(['@', '$#,##0.00'])])
        .setValues(invoiceHeaderValues) // Header Values

      if (n != numPages - 2)
        packingSlip.getRange(numRowsPerPage + 9 + N, 1, numItemsPerPage + 1, col) // Item information and formatting
          .setBorder(true, true, true, true, true, false).setFontColor('black').setFontFamily('Arial')
          .setHorizontalAlignments(new Array(numItemsPerPage + 1).fill(['center', 'center', 'left', 'left', 'left', 'left', 'left', 'right', 'right']))
          .setVerticalAlignments(new Array(numItemsPerPage + 1).fill(new Array(col).fill('middle')))
          .setFontSizes([new Array(col).fill(12), ...new Array(numItemsPerPage).fill(new Array(col).fill(9))])
          .setFontWeights([new Array(col).fill('bold'), ...new Array(numItemsPerPage).fill(new Array(col).fill('normal'))])
          .setNumberFormats([new Array(col).fill('@'), ...new Array(numItemsPerPage).fill(['@', '@', '@', '@', '@', '@', '@', '$#,##0.00', '$#,##0.00'])])
          .setValues([['Qty', 'SKU', 'Item', null, null, null, null, 'Price', 'Total'], ...itemInformation.slice(numItemsOnPageOne + numItemsPerPage*n, numItemsOnPageOne + numItemsPerPage*(n + 1))])
      else // Last Page
        packingSlip.getRange(numRowsPerPage + 9 + N, 1, numItemsPerPage + 1, col) // Item information and formatting
          .setBorder(true, true, true, true, true, false).setFontColor('black').setFontFamily('Arial')
          .setHorizontalAlignments(new Array(numItemsPerPage + 1).fill(['center', 'center', 'left', 'left', 'left', 'left', 'left', 'right', 'right']))
          .setVerticalAlignments(new Array(numItemsPerPage + 1).fill(new Array(col).fill('middle')))
          .setFontSizes([new Array(col).fill(12), ...new Array(numItemsPerPage).fill(new Array(col).fill(9))])
          .setFontWeights([new Array(col).fill('bold'), ...new Array(numItemsPerPage).fill(new Array(col).fill('normal'))])
          .setNumberFormats([new Array(col).fill('@'), ...new Array(numItemsPerPage).fill(['@', '@', '@', '@', '@', '@', '@', '$#,##0.00', '$#,##0.00'])])
          .setValues([['Qty', 'SKU', 'Item', null, null, null, null, 'Price', 'Total'], 
            ...itemInformation.slice(numItemsOnPageOne + numItemsPerPage*n, itemInformation.length), 
            ...new Array(numItemsOnPageOne + numItemsPerPage*(n + 1) - itemInformation.length).fill(new Array(col).fill(''))])
    }

    // setValues of total
    subtotalAmount += ')';
    packingSlip.getRange(3, col).setFormula(subtotalAmount)
  }
  else
  {
    packingSlip.getRange(16, 1, numItemsOnPageOne + 1, col) // Item information and formatting
      .setBorder(true, true, true, true, true, false).setFontColor('black').setFontFamily('Arial')
      .setHorizontalAlignments(new Array(numItemsOnPageOne + 1).fill(['center', 'center', 'left', 'left', 'left', 'left', 'left', 'right', 'right']))
      .setVerticalAlignments(new Array(numItemsOnPageOne + 1).fill(new Array(col).fill('middle')))
      .setFontSizes([new Array(col).fill(12), ...new Array(numItemsOnPageOne).fill(new Array(col).fill(9))])
      .setFontWeights([new Array(col).fill('bold'), ...new Array(numItemsOnPageOne).fill(new Array(col).fill('normal'))])
      .setNumberFormats([new Array(col).fill('@'), ...new Array(numItemsOnPageOne).fill(['@', '@', '@', '@', '@', '@', '@', '$#,##0.00', '$#,##0.00'])])  
      .setValues([['Qty', 'SKU', 'Item', null, null, null, null, 'Price', 'Total'],
        ...itemInformation.slice(0, itemInformation.length), 
        ...new Array(numItemsOnPageOne - itemInformation.length).fill(new Array(col).fill(''))])

    if (packingSlip.getMaxRows() > numRowsPerPage + 1)
      packingSlip.deleteRows(numRowsPerPage + 2, packingSlip.getMaxRows() - numRowsPerPage - 1) // One page, delete the extra rows
      
    packingSlip.getRange(numRowsPerPage + 1, col).setValue('') // Set the page number blank

    var namedRange = '';

    packingSlip.getNamedRanges().map(rng => {
      namedRange = rng.getName()

      if (namedRange[namedRange.length - 1] !== '1' && !isNaN(parseInt(namedRange[namedRange.length - 1]))) // If the range ends with a number that is not 1, then remove it
        rng.remove();
    })

    packingSlip.getRange(3, col).setFormula('=SUM(Item_Totals_Page_1)')
  }

  checkboxRange.setValues(checks)
}

/**
 * This function updates the Status Page by checking the shopify data for new orders that don't already exist on the Status page and then
 * places them there in descending order with respect to order number.
 * 
 * @param {Object[][]}  shopifyData : The data downloaded from shopify.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @return {Number} Returns the number of orders in the particular shopify data set.
 * @author Jarren Ralf
 */
function updateStatusPage(shopifyData, spreadsheet)
{
  const statusPage = spreadsheet.getSheetByName('Status Page');
  const statusPageData = statusPage.getSheetValues(3, 1, statusPage.getLastRow() - 2, statusPage.getLastColumn());
  var numAdditionalOrders = 0;

  for (var i = 1; i < shopifyData.length; i++)
  {
    if (shopifyData[i][0] != shopifyData[i - 1][0]) // New order or first order
    {
      for (var j = 0; j < statusPageData.length; j++)
      {
        if (shopifyData[i][0] === statusPageData[j][0]) // The status page already contains this order, so skip it
          break;
          
        if (j === statusPageData.length - 1) 
          statusPageData.push([shopifyData[i][0], // Order #,
            "FALSE",
            toProper(shopifyData[i][24]), // Billing Name
            null,
            isBlank(shopifyData[i][39]) ?          toProper(shopifyData[i][29]) : toProper(shopifyData[i][39]), // Shipping City or Billing City (if Pick Up)
            isBlank(shopifyData[i][41]) ?      shopifyData[i][31].toUpperCase() : shopifyData[i][41].toUpperCase(), // Shipping Province or Billing Province (if Pick Up)
            isBlank(shopifyData[i][43]) ? formatPhoneNumber(shopifyData[i][33]) : formatPhoneNumber(shopifyData[i][43]), // Shipping Phone Number or Billing Phone Number (if Pick Up)
            shopifyData[i][ 1].toString().toLowerCase(), // Email
            shopifyData[i][ 2], // Financial Status
            shopifyData[i][ 4], // Fulfillment Status
            shopifyData[i][14], // Shipping Method
            shopifyData[i][15].split(' ', 1)[0] // Created At
          ])
      }

      numAdditionalOrders++;
    }
  }

  const numOrders = statusPageData.length;
  const numCols = statusPageData[0].length;
  const numberFormats = new Array(numOrders).fill(['@', '#', ...new Array(numCols - 3).fill('@'), "dd MMM yyyy"])

  statusPageData.sort((a, b) => a[0] < b[0] ? 1 : -1) // Place the data in descending order with respect to order number

  statusPage.activate().getRange(3, 1, numOrders, numCols).setNumberFormats(numberFormats).setValues(statusPageData)

  return numAdditionalOrders;
}