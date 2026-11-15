function strategoselectpiecered() {
  const redboard = SpreadsheetApp.getActive().getSheetByName("Red's Stratego Board")
  const blueboard = SpreadsheetApp.getActive().getSheetByName("Blue's Stratego Board")

  //Make sure that it is your turn
  if (redboard.getRange('M5').getValue()!='Move it Where?' && redboard.getRange('M5').getValue()!='Your Turn') {
    Browser.msgBox('It is not your turn', 'You must wait for your opponent to make their move first', Browser.Buttons.OK)
    return
  }

  //Remove the previously selcted piece if there is one
  if (redboard.getRange('M5').getValue()=='Move it Where?') {
    var column = 0
    var row = 0
    var colors = redboard.getRange('B2:K9').getBackgrounds()
    for (i=0; i<8; i++) {
      for (j=0; j<10; j++) {
        if (colors[i][j]=='#ff0000') {
          row = i+2
          column = j+2
          break
        }
      }
      if (row>0 && column>0) {break}
    }
    redboard.getRange(row, column).setBackground('#ea9999')
    blueboard.getRange(row, column).setBackground('#ea9999')
  }

  var selectedpiece = redboard.getActiveCell()
  var piecetype = selectedpiece.getValue()

  //Handle other errors
  if ((selectedpiece.getBackground()!='#ea9999' && selectedpiece.getBackground()!='#ff0000') || selectedpiece.getColumn()==1 || selectedpiece.getColumn()>11 || selectedpiece.getRow()==1 || selectedpiece.getRow()>9) {
    Browser.msgBox('That is not one of your pieces', 'Please select one of your pieces to move', Browser.Buttons.OK)
    return
  }

  if (piecetype=='b' || piecetype=='B') {
    Browser.msgBox('Bombs cannot be moved', 'Please select a different piece to move', Browser.Buttons.OK)
    return
  }

  if (piecetype=='f' || piecetype=='F') {
    Browser.msgBox('Flags cannot be moved', 'Please select a different piece to move', Browser.Buttons.OK)
    return
  }

  if (piecetype=='s') {piecetype='S'}

  //Secretly store the information about the selcted piece in the top left of the sheet
  redboard.getRange('A1').setFontColor(redboard.getRange('A1').getBackground()).setValue(selectedpiece.getRow()+','+selectedpiece.getColumn())
  redboard.getRange('A2').setFontColor(redboard.getRange('A2').getBackground()).setValue(piecetype)

  //Color the cell on both boards and ask the red player where they want that piece moved to
  selectedpiece.setBackground('#ff0000')
  blueboard.getRange(selectedpiece.getRow(), selectedpiece.getColumn()).setBackground('#ff0000')
  redboard.getRange('M5').setValue('Move it Where?')
}
