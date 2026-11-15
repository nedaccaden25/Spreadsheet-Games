function flag() {
  
  const sheet = SpreadsheetApp.getActive().getSheetByName('Minesweeper')
  //Gets which page in the spreadsheet that these changes will be done on

  var toflag = sheet.getActiveCell()
  //Stores the highlighted cell

  var currentcolor = toflag.getBackground()
  //Store the color of the highlighted cell

  if (currentcolor == '#999999' || currentcolor == '#000000' || toflag.getRow() > 20 || !sheet.getRange('M21').isBlank()) {
    return
  }
  //If this space has been dug up, is off the board, or if the game is over, do nothing

  if (currentcolor == '#ffffff') {
    toflag.setBackground('#ea9999')
  }
  //If this space has no mine, place a flag with a normal color

  if (currentcolor == '#fffffe') {
    toflag.setBackground('#ea9998') 
  }
  //If this space has a mine, place a flag with a slightly different color

  if (currentcolor == '#ea9999') {
    toflag.setBackground('#ffffff')
  }
  //If this space has a flag and no mine, reset the space to white

  if (currentcolor == '#ea9998') {
    toflag.setBackground('#fffffe')
  }
  //If this space has a flag and a mine, reset the space to slightly of white
}
