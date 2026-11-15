function strategoselectpieceblue() {
  const blueboard = SpreadsheetApp.getActive().getSheetByName("Blue's Stratego Board")
  const redboard = SpreadsheetApp.getActive().getSheetByName("Red's Stratego Board")

  //Make sure that it is your turn
  if (blueboard.getRange('M5').getValue()!='Move it Where?' && blueboard.getRange('M5').getValue()!='Your Turn') {
    Browser.msgBox('It is not your turn', 'You must wait for your opponent to make their move first', Browser.Buttons.OK)
    return
  }

  //Remove the previously selcted piece if there is one
  if (blueboard.getRange('M5').getValue()=='Move it Where?') {
    var column = 0
    var row = 0
    var colors = blueboard.getRange('B2:K9').getBackgrounds()
    for (i=0; i<8; i++) {
      for (j=0; j<10; j++) {
        if (colors[i][j]=='#0000ff') {
          row = i+2
          column = j+2
          break
        }
      }
      if (row>0 && column>0) {break}
    }
    blueboard.getRange(row, column).setBackground('#a4c2f4')
    redboard.getRange(row, column).setBackground('#a4c2f4')
  }

  var selectedpiece = blueboard.getActiveCell()
  var piecetype = selectedpiece.getValue()

  //Handle other errors
  if ((selectedpiece.getBackground()!='#a4c2f4' && selectedpiece.getBackground()!='#0000ff') || selectedpiece.getColumn()==1 || selectedpiece.getColumn()>11 || selectedpiece.getRow()==1 || selectedpiece.getRow()>9) {
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
  blueboard.getRange('A1').setFontColor(redboard.getRange('A1').getBackground()).setValue(selectedpiece.getRow()+','+selectedpiece.getColumn())
  blueboard.getRange('A2').setFontColor(redboard.getRange('A2').getBackground()).setValue(piecetype)

  //Color the cell on both boards and ask the blue player where they want that piece moved to
  selectedpiece.setBackground('#0000ff')
  redboard.getRange(selectedpiece.getRow(), selectedpiece.getColumn()).setBackground('#0000ff')
  blueboard.getRange('M5').setValue('Move it Where?')
}
