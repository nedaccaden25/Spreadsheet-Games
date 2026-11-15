function checkers() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Checkers')

  var resetyn = Browser.msgBox('Are you sure you want to start a new game?', '', Browser.Buttons.YES_NO)

  if (resetyn == 'yes') {

    sheet.getRange('A1:H8').setBackground('#cccccc').setFontSize(8).setFontColor('#000000').removeCheckboxes()
    sheet.getRange('A1').setBackground('#a4c2f4').insertCheckboxes()
    sheet.getRange('C1').setBackground('#a4c2f4').insertCheckboxes()
    sheet.getRange('E1').setBackground('#a4c2f4').insertCheckboxes()
    sheet.getRange('G1').setBackground('#a4c2f4').insertCheckboxes()
    sheet.getRange('B2').setBackground('#a4c2f4').insertCheckboxes()
    sheet.getRange('D2').setBackground('#a4c2f4').insertCheckboxes()
    sheet.getRange('F2').setBackground('#a4c2f4').insertCheckboxes()
    sheet.getRange('H2').setBackground('#a4c2f4').insertCheckboxes()
    sheet.getRange('A3').setBackground('#a4c2f4').insertCheckboxes()
    sheet.getRange('C3').setBackground('#a4c2f4').insertCheckboxes()
    sheet.getRange('E3').setBackground('#a4c2f4').insertCheckboxes()
    sheet.getRange('G3').setBackground('#a4c2f4').insertCheckboxes()
    sheet.getRange('B4').setBackground('#ffffff').insertCheckboxes()
    sheet.getRange('D4').setBackground('#ffffff').insertCheckboxes()
    sheet.getRange('F4').setBackground('#ffffff').insertCheckboxes()
    sheet.getRange('H4').setBackground('#ffffff').insertCheckboxes()
    sheet.getRange('A5').setBackground('#ffffff').insertCheckboxes()
    sheet.getRange('C5').setBackground('#ffffff').insertCheckboxes()
    sheet.getRange('E5').setBackground('#ffffff').insertCheckboxes()
    sheet.getRange('G5').setBackground('#ffffff').insertCheckboxes()
    sheet.getRange('B6').setBackground('#ea9999').insertCheckboxes()
    sheet.getRange('D6').setBackground('#ea9999').insertCheckboxes()
    sheet.getRange('F6').setBackground('#ea9999').insertCheckboxes()
    sheet.getRange('H6').setBackground('#ea9999').insertCheckboxes()
    sheet.getRange('A7').setBackground('#ea9999').insertCheckboxes()
    sheet.getRange('C7').setBackground('#ea9999').insertCheckboxes()
    sheet.getRange('E7').setBackground('#ea9999').insertCheckboxes()
    sheet.getRange('G7').setBackground('#ea9999').insertCheckboxes()
    sheet.getRange('B8').setBackground('#ea9999').insertCheckboxes()
    sheet.getRange('D8').setBackground('#ea9999').insertCheckboxes()
    sheet.getRange('F8').setBackground('#ea9999').insertCheckboxes()
    sheet.getRange('H8').setBackground('#ea9999').insertCheckboxes()

    sheet.getRange('I1').setBackground('#ea9999').setValue("Red's turn").setFontColor('#000000').setFontSize(20)
  }
}
