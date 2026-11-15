function mineGame() {
  
  const sheet = SpreadsheetApp.getActive().getSheetByName('Minesweeper')
  //Gets which page in the spreadsheet that these changes will be done on

  var resetyn = Browser.msgBox('Are you sure you want to make a new board?', '', Browser.Buttons.YES_NO)
  //Makes the pop-up

  if (resetyn === "yes") {
  //Checks which option you selected, the rest of the code will only run if yes was selected

    var nummines = sheet.getRange('G24').getValue()
    //Store how many mines need to be placed

    var minelist = new Array(nummines)
    //Generates an array to store the locations of all the mines

    var boardheight = sheet.getRange('A24').getValue()
    //Store the height of the board to generate

    if (boardheight == '') {
      boardheight = Math.floor(Math.random() * (13))+8
    }
    //If no value was inputted, fill with a random value

    var boardwidth = sheet.getRange('D24').getValue()
    //Store the width of the board to generate

    if (boardwidth == '') {
      boardwidth = Math.floor(Math.random() * (13))+8
    }
    //If no value was inputted, fill with a random value

    if (nummines == '') {
      nummines = Math.floor(Math.random() * (boardheight*boardwidth-(boardheight+boardwidth)))
      if (nummines/(boardheight*boardwidth)>0.5) {
        nummines = Math.floor(nummines/3)
      }
    }
    //If no value was inputted, fill with a random value based on the size of the board
    //Limits the randomness if it covers more than half of the board
    //These operations were just chosen by feel, no real method here other than trying to keep it fair

    var temppos = 0
    //Temporary position to hold a mine position before confirming it

    var tempposfailed = 0
    //Flag used for determining if a mine position is a duplicate

    if (boardheight > 20 || boardheight < 2 || boardwidth > 20 || boardwidth < 2) {

      Browser.msgBox("Could not generate board", "Invalid board size", Browser.Buttons.OK)
      return
      //Checks for invalid board size
    }

    if (nummines < 1 || nummines >= boardheight*boardwidth) {

      Browser.msgBox("Could not generate board", "Invalid number of mines", Browser.Buttons.OK)
      return
      //Checks for invalid number of mines
    }

    sheet.getRange('A1:T20').setValue('').setBackground('#999999').setFontColor('#000000')
    //Color the entire board gray, clear all text, and set the font color to black

    SpreadsheetApp.flush()
    //Applies the changes

    sheet.getRange('M21').setValue('').setBackground('#ffffff')
    //Clear result box

    if (boardheight != 20) {
      sheet.getRange(boardheight+1, 1, 20-boardheight, 20).setBackground('#000000')
      //Limit the board height by coloring unused spaces black
    }

    if (boardwidth != 20) {
      sheet.getRange(1, boardwidth+1, 20, 20-boardwidth).setBackground('#000000')
      //Limit the board width by coloring unused spaces black
    }

    var i = nummines
    //Counter varible used for placing the mines

    while (i!=0) {
      temppos = sheet.getRange(Math.floor(Math.random() * (boardheight-1))+1, Math.floor(Math.random() * (boardwidth-1))+1)
      //Generate a random position on the board

      for (j=nummines-i; j!=0; j--) {

        if (minelist[j-1].getColumn()==temppos.getColumn() && minelist[j-1].getRow()==temppos.getRow()) {
          tempposfailed = 1
          break
        }
      }
      //Check if that position was already claimed

      if (tempposfailed == 0) {
        minelist[nummines-i] = temppos
        i=i-1
      }
      //If the position is not a duplicate, decrement the counter and add the position to the list

      tempposfailed = 0
      //Reset the flag before restarting the loop

    }

    sheet.getRange(1,1, boardheight, boardwidth).setBackground('#ffffff')
    //Sets the board color to white
    //This wasn't done initially in order to make it clear that the board is not done generating while it is gray

    for (i=nummines; i!=0; i--) {
      minelist[i-1].setBackground('#fffffe')
    }
    //Color all of the mines a slightly different color

  }
}
