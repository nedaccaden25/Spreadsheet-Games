function checkersforfeit() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Checkers')

  var resetyn = Browser.msgBox('Are you sure you want to forfeit the game?', '', Browser.Buttons.YES_NO)

  if (resetyn == 'yes') {

    if (sheet.getRange('I1').getValue()=="Red's turn") {
      sheet.getRange('I1').setValue('Blue Wins!').setBackground('#a4c2f4')
    }
    if (sheet.getRange('I1').getValue()=="Blue's turn") {
      sheet.getRange('I1').setValue('Red Wins!').setBackground('#ea9999')
    }
  }
}
