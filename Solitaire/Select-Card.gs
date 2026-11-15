function solitaireselectcard() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Solitaire')

  var selectedcell = sheet.getActiveCell()
  var selectedcolor = selectedcell.getBackground()
  var boardcolors = sheet.getRange('A1:O27').getBackgrounds()

  if (selectedcolor=='#ea9999') {return}

  //Handle errors
  if (selectedcolor=='#d9ead3') {
    Browser.msgBox('Please select a valid card to move', '', Browser.Buttons.OK)
    return
  }

  if (selectedcolor=='#a4c2f4') {
    Browser.msgBox('You cannot move unrevealed cards or cards from the stockpile', '', Browser.Buttons.OK)
    return
  }

  if (selectedcell.getRow()>27 || selectedcell.getColumn()>15) {
    Browser.msgBox('Please select a card from the board to move', '', Browser.Buttons.OK)
    return
  }

  //Deselect a previously selected piece if there is one
  for (i=0; i<27; i++) {
    for (j=0; j<15; j++) {
      if (boardcolors[i][j]=='#ea9999') {
        sheet.getRange('A1').offset(i, j).setBackground('#ffffff')
      }
    }
  }

  //Color the selected cells
  if (selectedcell.getRow()>3) {
    for (i=0; i>-1; i++) {
      if (boardcolors[selectedcell.getRow()-1+i][selectedcell.getColumn()-1]=='#ffffff') {
        selectedcell.offset(i, 0).setBackground('#ea9999')
      } else {break}
    }
  } else {
    if (selectedcell.offset(3, 0).getBackground()=='#ffffff') {
      Browser.msgBox('You can only select the top card drawn to move', '', Browser.Buttons.OK)
      return
    } else {
      selectedcell.setBackground('#ea9999')
    }
  }
}
