function yahtzeeroll() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Yahtzee')
  var r = 0

  if (sheet.getRange('D1').getValue()<=2 &&(sheet.getRange('Z10').getValue()+sheet.getRange('AA10').getValue()+sheet.getRange('AB10').getValue()+sheet.getRange('AC10').getValue()!=0)){
    if (sheet.getRange('Z10').getValue()=='X1') { if (!sheet.getRange('AC22').isBlank()) {return}}

    for (j=0; j!=5; j++) {
      if (!sheet.getRange('B6').offset(0, j*4).isChecked() || sheet.getRange('D1').getValue()==0) {
        r = Math.floor(Math.random()*6+1)
        
        sheet.getRange('B2:D4').offset(0, j*4).clearContent()
        SpreadsheetApp.flush()

        if (r == 1) {sheet.getRange('C3').offset(0, j*4).setValue('•')}

        if (r == 2) {sheet.getRange('B4').offset(0, j*4).setValue('•')
          sheet.getRange('D2').offset(0, j*4).setValue('•')}

        if (r == 3) {sheet.getRange('B4').offset(0, j*4).setValue('•')
          sheet.getRange('C3').offset(0, j*4).setValue('•')
          sheet.getRange('D2').offset(0, j*4).setValue('•')}

        if (r == 4) {sheet.getRange('B2').offset(0, j*4).setValue('•')
          sheet.getRange('B4').offset(0, j*4).setValue('•')
          sheet.getRange('D2').offset(0, j*4).setValue('•')
          sheet.getRange('D4').offset(0, j*4).setValue('•')}

        if (r == 5) {sheet.getRange('C3').offset(0, j*4).setValue('•')
          sheet.getRange('B2').offset(0, j*4).setValue('•')
          sheet.getRange('B4').offset(0, j*4).setValue('•')
          sheet.getRange('D2').offset(0, j*4).setValue('•')
          sheet.getRange('D4').offset(0, j*4).setValue('•')}

        if (r == 6) {sheet.getRange('B2').offset(0, j*4).setValue('•')
          sheet.getRange('B3').offset(0, j*4).setValue('•')
          sheet.getRange('B4').offset(0, j*4).setValue('•')
          sheet.getRange('D2').offset(0, j*4).setValue('•')
          sheet.getRange('D3').offset(0, j*4).setValue('•')
          sheet.getRange('D4').offset(0, j*4).setValue('•')}
      }
    }
    sheet.getRange('D1').setValue(sheet.getRange('D1').getValue()+1)
  }
}
