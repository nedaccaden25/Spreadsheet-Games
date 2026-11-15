function dig() {
  
  const sheet = SpreadsheetApp.getActive().getSheetByName('Minesweeper')
  //Gets which page in the spreadsheet that these changes will be done on

  var todig = sheet.getActiveCell()
  //Stores the highlighted cell

  var mode = 0
  //A flag used for checking if adjacent spaces have already been dealt with

  var numadjmines = 0
  //Variable that keeps track of how many mines are near this space

  var boardheight = 0
  //Keeps track of the board height

  var boardwidth = 0
  //Keeps track of the board width

  for (i=20; i!=0; i--) {
    if (sheet.getRange(i, 1,).getBackground()!='#000000') {
      boardheight = i
      break
    }
  }
  //Gets the height of the board based on the black space

  for (i=20; i!=0; i--) {
    if (sheet.getRange(1, i).getBackground()!='#000000') {
      boardwidth = i
      break
    }
  }
  //Gets the width of the board based on the black space

  if (todig.getBackground()=='#000000' || todig.getRow()>20 || !sheet.getRange('M21').isBlank()) {
    return
  }
  //If this cell is not on the board, has already been cleared, or if the game is over, don't do anything

  if (todig.getBackground()=='#fffffe') {
    todig.setBackground('#ff0000')
    sheet.getRange('M21').setValue('KABOOM!!!').setBackground('#ff0000')
    return
  }
  //If this cell was a mine, blow it up and display the KABOOM!!!

  if (todig.getColumn()==1 && todig.getRow()==1) {mode = 'topLeft'}
  //This is the top left cell check the appropriate cells

  if (mode===0 && (todig.getColumn()==boardwidth && todig.getRow()==1)) {mode = 'topRight'}
  //This is the top right cell check the appropriate cells

  if (mode===0 && (todig.getColumn()==1 && todig.getRow()==boardheight)) {mode = 'bottomLeft'}
  //This is the bottom left cell check the appropriate cells

  if (mode===0 && (todig.getColumn()==boardwidth && todig.getRow()==boardheight)) {mode = 'bottomRight'}
  //This is the bottom right cell check the appropriate cells

  if (mode===0 && todig.getRow()==1) {mode = 'topRow'}
  //This is the top row check the appropriate cells

  if (mode===0 && todig.getColumn()==1) {mode = 'leftCol'}
  //This is the left column check the appropriate cells

  if (mode===0 && todig.getRow()==boardheight) {mode = 'bottomRow'}
  //This is the bottom row check the appropriate cells

  if (mode===0 && todig.getColumn()==boardwidth) {mode = 'rightCol'}
  //This is the right column check the appropriate cells

  if (mode===0) {mode = 'center'}
  //This is the center check the approprite cells

  if (todig.getBackground()=='#ffffff') {
    todig.setBackground('#999999')
    //Turn the cell gray

    switch(mode) {
      case 'topLeft': numadjmines = topLeft(todig, sheet, '#fffffe', '#ea9998')
        break
      case 'topRight': numadjmines = topRight(todig, sheet, '#fffffe', '#ea9998')
        break
      case 'bottomLeft': numadjmines = bottomLeft(todig, sheet, '#fffffe', '#ea9998')
        break
      case 'bottomRight': numadjmines = bottomRight(todig, sheet, '#fffffe', '#ea9998')
        break
      case 'topRow': numadjmines = topRow(todig, sheet, '#fffffe', '#ea9998')
        break
      case 'leftCol': numadjmines = leftCol(todig, sheet, '#fffffe', '#ea9998')
        break
      case 'bottomRow': numadjmines = bottomRow(todig, sheet, '#fffffe', '#ea9998')
        break
      case 'rightCol': numadjmines = rightCol(todig, sheet, '#fffffe', '#ea9998')
        break
      case 'center': numadjmines = center(todig, sheet, '#fffffe', '#ea9998')
        break
    }
    //Determine how many mines are next to this cell based on its location

    switch(numadjmines) {
      case 1: todig.setValue(1).setFontColor('#1155cc')
        break
      case 2: todig.setValue(2).setFontColor('#38761d')
        break
      case 3: todig.setValue(3).setFontColor('#cc0000')
        break
      case 4: todig.setValue(4).setFontColor('#271cbe')
        break
      case 5: todig.setValue(5).setFontColor('#a61c00')
        break
      case 6: todig.setValue(6).setFontColor('#45818e')
        break
      case 7: todig.setValue(7).setFontColor('#000000')
        break
      case 8: todig.setValue(8).setFontColor('#666666')
        break
      case 9: todig.setValue('what')
        break
        //If this cell has at least one mine next to it, display that number

      case 0: 
        switch(mode) {
          case 'topLeft':
            if (todig.offset(0, 1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
            if (todig.offset(1, 0).getBackground() != '#999999') {recursion(todig.offset(1, 0), sheet, boardheight, boardwidth)}
            if (todig.offset(1, 1).getBackground() != '#999999') {recursion(todig.offset(1, 1), sheet, boardheight, boardwidth)}
            break

          case 'topRight':
            if (todig.offset(0, -1).getBackground() != '#999999') {recursion(todig.offset(0, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(1, -1).getBackground() != '#999999') {recursion(todig.offset(1, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(1, 0).getBackground() != '#999999') {recursion(todig.offset(1, 0), sheet, boardheight, boardwidth)}
            break

          case 'bottomLeft':
            if (todig.offset(-1, 0).getBackground() != '#999999') {recursion(todig.offset(-1, 0), sheet, boardheight, boardwidth)}
            if (todig.offset(-1, 1).getBackground() != '#999999') {recursion(todig.offset(-1, 1), sheet, boardheight, boardwidth)}
            if (todig.offset(0, 1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
            break

          case 'bottomRight':
            if (todig.offset(-1, -1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
            if (todig.offset(-1, 0).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
            if (todig.offset(0, -1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
            break

          case 'topRow':
            if (todig.offset(0, -1).getBackground() != '#999999') {recursion(todig.offset(0, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(0, 1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
            if (todig.offset(1, -1).getBackground() != '#999999') {recursion(todig.offset(1, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(1, 0).getBackground() != '#999999') {recursion(todig.offset(1, 0), sheet, boardheight, boardwidth)}
            if (todig.offset(1, 1).getBackground() != '#999999') {recursion(todig.offset(1, 1), sheet, boardheight, boardwidth)}
            break

          case 'leftCol':
            if (todig.offset(-1, 0).getBackground() != '#999999') {recursion(todig.offset(-1, 0), sheet, boardheight, boardwidth)}
            if (todig.offset(-1, 1).getBackground() != '#999999') {recursion(todig.offset(-1, 1), sheet, boardheight, boardwidth)}
            if (todig.offset(0, 1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
            if (todig.offset(1, 0).getBackground() != '#999999') {recursion(todig.offset(1, 0), sheet, boardheight, boardwidth)}
            if (todig.offset(1, 1).getBackground() != '#999999') {recursion(todig.offset(1, 1), sheet, boardheight, boardwidth)}
            break

          case 'bottomRow':
            if (todig.offset(-1, -1).getBackground() != '#999999') {recursion(todig.offset(-1, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(-1, 0).getBackground() != '#999999') {recursion(todig.offset(-1, 0), sheet, boardheight, boardwidth)}
            if (todig.offset(-1, 1).getBackground() != '#999999') {recursion(todig.offset(-1, 1), sheet, boardheight, boardwidth)}
            if (todig.offset(0, -1).getBackground() != '#999999') {recursion(todig.offset(0, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(0, 1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
            break

          case 'rightCol':
            if (todig.offset(-1, -1).getBackground() != '#999999') {recursion(todig.offset(-1, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(-1, 0).getBackground() != '#999999') {recursion(todig.offset(-1, 0), sheet, boardheight, boardwidth)}
            if (todig.offset(0, -1).getBackground() != '#999999') {recursion(todig.offset(0, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(1, -1).getBackground() != '#999999') {recursion(todig.offset(1, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(1, 0).getBackground() != '#999999') {recursion(todig.offset(1, 0), sheet, boardheight, boardwidth)}
            break

          case 'center':
            if (todig.offset(-1, -1).getBackground() != '#999999') {recursion(todig.offset(-1, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(-1, 0).getBackground() != '#999999') {recursion(todig.offset(-1, 0), sheet, boardheight, boardwidth)}
            if (todig.offset(-1, 1).getBackground() != '#999999') {recursion(todig.offset(-1, 1), sheet, boardheight, boardwidth)}
            if (todig.offset(0, -1).getBackground() != '#999999') {recursion(todig.offset(0, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(0, 1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
            if (todig.offset(1, -1).getBackground() != '#999999') {recursion(todig.offset(1, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(1, 0).getBackground() != '#999999') {recursion(todig.offset(1, 0), sheet, boardheight, boardwidth)}
            if (todig.offset(1, 1).getBackground() != '#999999') {recursion(todig.offset(1, 1), sheet, boardheight, boardwidth)}
            break

          default: return
        }
        break
        //If this cell has no mines next to it, perform this same function with each neighboring cell that hasn't been dug up
      default: return
    }
  }

  if (todig.getBackground()=='#999999' && !todig.isBlank()) {
  //If this is instead a cell that has already been dug up and has a number, possibly run a shortcut

    switch(mode) {
      case 'topLeft': numadjmines = topLeft(todig, sheet, '#ea9999', '#ea9998')
        break
      case 'topRight': numadjmines = topRight(todig, sheet, '#ea9999', '#ea9998')
        break
      case 'bottomLeft': numadjmines = bottomLeft(todig, sheet, '#ea9999', '#ea9998')
        break
      case 'bottomRight': numadjmines = bottomRight(todig, sheet, '#ea9999', '#ea9998')
        break
      case 'topRow': numadjmines = topRow(todig, sheet, '#ea9999', '#ea9998')
        break
      case 'leftCol': numadjmines = leftCol(todig, sheet, '#ea9999', '#ea9998')
        break
      case 'bottomRow': numadjmines = bottomRow(todig, sheet, '#ea9999', '#ea9998')
        break
      case 'rightCol': numadjmines = rightCol(todig, sheet, '#ea9999', '#ea9998')
        break
      case 'center': numadjmines = center(todig, sheet, '#ea9999', '#ea9998')
        break
    }
    //Determine how many flags are next to this cell based on its location

    if (numadjmines == todig.getValue()) {
      switch(mode) {
        case 'topLeft':
          if (todig.offset(0, 1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidthe)}
          if (todig.offset(1, 0).getBackground() != '#999999') {recursion(todig.offset(1, 0), sheet, boardheight, boardwidth)}
          if (todig.offset(1, 1).getBackground() != '#999999') {recursion(todig.offset(1, 1), sheet, boardheight, boardwidth)}
          break

        case 'topRight':
          if (todig.offset(0, -1).getBackground() != '#999999') {recursion(todig.offset(0, -1), sheet, boardheight, boardwidth)}
          if (todig.offset(1, -1).getBackground() != '#999999') {recursion(todig.offset(1, -1), sheet, boardheight, boardwidth)}
          if (todig.offset(1, 0).getBackground() != '#999999') {recursion(todig.offset(1, 0), sheet, boardheight, boardwidth)}
          break

        case 'bottomLeft':
          if (todig.offset(-1, 0).getBackground() != '#999999') {recursion(todig.offset(-1, 0), sheet, boardheight, boardwidth)}
          if (todig.offset(-1, 1).getBackground() != '#999999') {recursion(todig.offset(-1, 1), sheet, boardheight, boardwidth)}
          if (todig.offset(0, 1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
          break

        case 'bottomRight':
          if (todig.offset(-1, -1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
          if (todig.offset(-1, 0).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
          if (todig.offset(0, -1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
          break

        case 'topRow':
          if (todig.offset(0, -1).getBackground() != '#999999') {recursion(todig.offset(0, -1), sheet, boardheight, boardwidth)}
          if (todig.offset(0, 1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
          if (todig.offset(1, -1).getBackground() != '#999999') {recursion(todig.offset(1, -1), sheet, boardheight, boardwidth)}
          if (todig.offset(1, 0).getBackground() != '#999999') {recursion(todig.offset(1, 0), sheet, boardheight, boardwidth)}
          if (todig.offset(1, 1).getBackground() != '#999999') {recursion(todig.offset(1, 1), sheet, boardheight, boardwidth)}
          break

        case 'leftCol':
          if (todig.offset(-1, 0).getBackground() != '#999999') {recursion(todig.offset(-1, 0), sheet, boardheight, boardwidth)}
          if (todig.offset(-1, 1).getBackground() != '#999999') {recursion(todig.offset(-1, 1), sheet, boardheight, boardwidth)}
          if (todig.offset(0, 1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
          if (todig.offset(1, 0).getBackground() != '#999999') {recursion(todig.offset(1, 0), sheet, boardheight, boardwidth)}
          if (todig.offset(1, 1).getBackground() != '#999999') {recursion(todig.offset(1, 1), sheet, boardheight, boardwidth)}
          break

        case 'bottomRow':
          if (todig.offset(-1, -1).getBackground() != '#999999') {recursion(todig.offset(-1, -1), sheet, boardheight, boardwidth)}
          if (todig.offset(-1, 0).getBackground() != '#999999') {recursion(todig.offset(-1, 0), sheet, boardheight, boardwidth)}
          if (todig.offset(-1, 1).getBackground() != '#999999') {recursion(todig.offset(-1, 1), sheet, boardheight, boardwidth)}
          if (todig.offset(0, -1).getBackground() != '#999999') {recursion(todig.offset(0, -1), sheet, boardheight, boardwidth)}
          if (todig.offset(0, 1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
          break

        case 'rightCol':
          if (todig.offset(-1, -1).getBackground() != '#999999') {recursion(todig.offset(-1, -1), sheet, boardheight, boardwidth)}
          if (todig.offset(-1, 0).getBackground() != '#999999') {recursion(todig.offset(-1, 0), sheet, boardheight, boardwidth)}
          if (todig.offset(0, -1).getBackground() != '#999999') {recursion(todig.offset(0, -1), sheet, boardheight, boardwidth)}
          if (todig.offset(1, -1).getBackground() != '#999999') {recursion(todig.offset(1, -1), sheet, boardheight, boardwidth)}
          if (todig.offset(1, 0).getBackground() != '#999999') {recursion(todig.offset(1, 0), sheet, boardheight, boardwidth)}
          break

        case 'center':
          if (todig.offset(-1, -1).getBackground() != '#999999') {recursion(todig.offset(-1, -1), sheet, boardheight, boardwidth)}
          if (todig.offset(-1, 0).getBackground() != '#999999') {recursion(todig.offset(-1, 0), sheet, boardheight, boardwidth)}
          if (todig.offset(-1, 1).getBackground() != '#999999') {recursion(todig.offset(-1, 1), sheet, boardheight, boardwidth)}
          if (todig.offset(0, -1).getBackground() != '#999999') {recursion(todig.offset(0, -1), sheet, boardheight, boardwidth)}
          if (todig.offset(0, 1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
          if (todig.offset(1, -1).getBackground() != '#999999') {recursion(todig.offset(1, -1), sheet, boardheight, boardwidth)}
          if (todig.offset(1, 0).getBackground() != '#999999') {recursion(todig.offset(1, 0), sheet, boardheight, boardwidth)}
          if (todig.offset(1, 1).getBackground() != '#999999') {recursion(todig.offset(1, 1), sheet, boardheight, boardwidth)}
          break

        default: return
      }
    }
  }

  var finalcolors = sheet.getRange(1, 1, boardheight, boardwidth).getBackgrounds()
  //Creates an array to store the colors of the entire board after all changes have been made

  var succeedflag = 0
  //Sets up a flag used for determining if you've won

  for (i=0; i!=boardheight; i++) {
    if (succeedflag == 1) {break}
    for (j=0; j!=boardwidth; j++) {
      if (succeedflag == 1) {break}
      if (finalcolors[i][j] == '#ffffff' || finalcolors[i][j] == '#ea9999') {
        succeedflag=1
        break
      }
    }
  }
  //This embedded loop checks each space to see if it is still buried or exploded

  if (succeedflag == 0) {
    sheet.getRange('M21').setValue('YOU WIN!').setBackground('#7fdf55')
  }
  //Display the winning message and color the result box green

}

function topLeft(todig, sheet, color1, color2) {

  var numadjmines = 0
  //Prep a counter to count the number of adjacent mines

  var adjspaces = sheet.getRange(todig.getRow(), todig.getColumn(), 2, 2).getBackgrounds()
  //Get the colors of all the cells surrounding the selected cell

  if (adjspaces[0][1] == color1 || adjspaces[0][1] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[1][0] == color1 || adjspaces[1][0] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[1][1] == color1 || adjspaces[1][1] == color2) {numadjmines=numadjmines+1}
  //Check the spaces for mines and adjust the counter as necessary

  return(numadjmines)
}

function topRight(todig, sheet, color1, color2) {

  var numadjmines = 0
  //Prep a counter to count the number of adjacent mines

  var adjspaces = sheet.getRange(todig.getRow(), todig.getColumn()-1, 2, 2).getBackgrounds()
  //Get the colors of all the cells surrounding the selected cell

  if (adjspaces[0][0] == color1 || adjspaces[0][0] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[1][0] == color1 || adjspaces[1][0] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[1][1] == color1 || adjspaces[1][1] == color2) {numadjmines=numadjmines+1}
  //Check the spaces for mines and adjust the counter as necessary

  return(numadjmines)
}

function bottomLeft(todig, sheet, color1, color2) {

  var numadjmines = 0
  //Prep a counter to count the number of adjacent mines

  var adjspaces = sheet.getRange(todig.getRow()-1, todig.getColumn(), 2, 2).getBackgrounds()
  //Get the colors of all the cells surrounding the selected cell

  if (adjspaces[0][0] == color1 || adjspaces[0][0] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[0][1] == color1 || adjspaces[0][1] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[1][1] == color1 || adjspaces[1][1] == color2) {numadjmines=numadjmines+1}
  //Check the spaces for mines and adjust the counter as necessary

  return(numadjmines)
}

function bottomRight(todig, sheet, color1, color2) {

  var numadjmines = 0
  //Prep a counter to count the number of adjacent mines

  var adjspaces = sheet.getRange(todig.getRow()-1, todig.getColumn()-1, 2, 2).getBackgrounds()
  //Get the colors of all the cells surrounding the selected cell

  if (adjspaces[0][0] == color1 || adjspaces[0][0] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[0][1] == color1 || adjspaces[0][1] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[1][0] == color1 || adjspaces[1][0] == color2) {numadjmines=numadjmines+1}
  //Check the spaces for mines and adjust the counter as necessary

  return(numadjmines)
}

function topRow(todig, sheet, color1, color2) {

  var numadjmines = 0
  //Prep a counter to count the number of adjacent mines

  var adjspaces = sheet.getRange(todig.getRow(), todig.getColumn()-1, 2, 3).getBackgrounds()
  //Get the colors of all the cells surrounding the selected cell

  if (adjspaces[0][0] == color1 || adjspaces[0][0] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[0][2] == color1 || adjspaces[0][2] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[1][0] == color1 || adjspaces[1][0] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[1][1] == color1 || adjspaces[1][1] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[1][2] == color1 || adjspaces[1][2] == color2) {numadjmines=numadjmines+1}
  //Check the spaces for mines and adjust the counter as necessary

  return(numadjmines)
}

function leftCol(todig, sheet, color1, color2) {

  var numadjmines = 0
  //Prep a counter to count the number of adjacent mines

  var adjspaces = sheet.getRange(todig.getRow()-1, todig.getColumn(), 3, 2).getBackgrounds()
  //Get the colors of all the cells surrounding the selected cell

  if (adjspaces[0][0] == color1 || adjspaces[0][0] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[0][1] == color1 || adjspaces[0][1] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[1][1] == color1 || adjspaces[1][1] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[2][0] == color1 || adjspaces[2][0] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[2][1] == color1 || adjspaces[2][1] == color2) {numadjmines=numadjmines+1}
  //Check the spaces for mines and adjust the counter as necessary

  return(numadjmines)
}

function bottomRow(todig, sheet, color1, color2) {

  var numadjmines = 0
  //Prep a counter to count the number of adjacent mines

  var adjspaces = sheet.getRange(todig.getRow()-1, todig.getColumn()-1, 2, 3).getBackgrounds()
  //Get the colors of all the cells surrounding the selected cell

  if (adjspaces[0][0] == color1 || adjspaces[0][0] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[0][1] == color1 || adjspaces[0][1] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[0][2] == color1 || adjspaces[0][2] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[1][0] == color1 || adjspaces[1][0] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[1][2] == color1 || adjspaces[1][2] == color2) {numadjmines=numadjmines+1}
  //Check the spaces for mines and adjust the counter as necessary

  return(numadjmines)
}

function rightCol(todig, sheet, color1, color2) {

  var numadjmines = 0
  //Prep a counter to count the number of adjacent mines

  var adjspaces = sheet.getRange(todig.getRow()-1, todig.getColumn()-1, 3, 2).getBackgrounds()
  //Get the colors of all the cells surrounding the selected cell

  if (adjspaces[0][0] == color1 || adjspaces[0][0] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[0][1] == color1 || adjspaces[0][1] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[1][0] == color1 || adjspaces[1][0] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[2][0] == color1 || adjspaces[2][0] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[2][1] == color1 || adjspaces[2][1] == color2) {numadjmines=numadjmines+1}
  //Check the spaces for mines and adjust the counter as necessary

  return(numadjmines)
}

function center(todig, sheet, color1, color2) {

  var numadjmines = 0
  //Prep a counter to count the number of adjacent mines

  var adjspaces = sheet.getRange(todig.getRow()-1, todig.getColumn()-1, 3, 3).getBackgrounds()
  //Get the colors of all the cells surrounding the selected cell

  if (adjspaces[0][0] == color1 || adjspaces[0][0] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[0][1] == color1 || adjspaces[0][1] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[0][2] == color1 || adjspaces[0][2] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[1][0] == color1 || adjspaces[1][0] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[1][2] == color1 || adjspaces[1][2] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[2][0] == color1 || adjspaces[2][0] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[2][1] == color1 || adjspaces[2][1] == color2) {numadjmines=numadjmines+1}
  if (adjspaces[2][2] == color1 || adjspaces[2][2] == color2) {numadjmines=numadjmines+1}
  //Check the spaces for mines and adjust the counter as necessary

  return(numadjmines)
}

function recursion(todig, sheet, boardheight, boardwidth) {

  if (sheet.getRange('M21').getValue() == 'KABOOM!!!') {
    return
  }
  //If a mine has been found, stop immediately

  var mode = 0
  //A flag used for checking if adjacent spaces have already been dealt with

  var numadjmines = 0
  //Variable that keeps track of how many mines are near this space

  if (todig.getBackground()=='#000000' || todig.getRow()>20) {
    return
  }
  //If this cell is not on the board don't do anything

  if (todig.getBackground()=='#fffffe') {
    todig.setBackground('#ff0000')
    sheet.getRange('M21').setValue('KABOOM!!!').setBackground('#ff0000')
    return
  }
  //If this cell was a mine, blow it up and display the KABOOM!!!

  if (todig.getColumn()==1 && todig.getRow()==1) {mode = 'topLeft'}
  //This is the top left cell check the appropriate cells

  if (mode===0 && (todig.getColumn()==boardwidth && todig.getRow()==1)) {mode = 'topRight'}
  //This is the top right cell check the appropriate cells

  if (mode===0 && (todig.getColumn()==1 && todig.getRow()==boardheight)) {mode = 'bottomLeft'}
  //This is the bottom left cell check the appropriate cells

  if (mode===0 && (todig.getColumn()==boardwidth && todig.getRow()==boardheight)) {mode = 'bottomRight'}
  //This is the bottom right cell check the appropriate cells

  if (mode===0 && todig.getRow()==1) {mode = 'topRow'}
  //This is the top row check the appropriate cells

  if (mode===0 && todig.getColumn()==1) {mode = 'leftCol'}
  //This is the left column check the appropriate cells

  if (mode===0 && todig.getRow()==boardheight) {mode = 'bottomRow'}
  //This is the bottom row check the appropriate cells

  if (mode===0 && todig.getColumn()==boardwidth) {mode = 'rightCol'}
  //This is the right column check the appropriate cells

  if (mode===0) {mode = 'center'}
  //This is the center check the approprite cells

  if (todig.getBackground() == '#ffffff') {
    todig.setBackground('#999999')
    //Turn the cell gray

    switch(mode) {
      case 'topLeft': numadjmines = topLeft(todig, sheet, '#fffffe', '#ea9998')
        break
      case 'topRight': numadjmines = topRight(todig, sheet, '#fffffe', '#ea9998')
        break
      case 'bottomLeft': numadjmines = bottomLeft(todig, sheet, '#fffffe', '#ea9998')
        break
      case 'bottomRight': numadjmines = bottomRight(todig, sheet, '#fffffe', '#ea9998')
        break
      case 'topRow': numadjmines = topRow(todig, sheet, '#fffffe', '#ea9998')
        break
      case 'leftCol': numadjmines = leftCol(todig, sheet, '#fffffe', '#ea9998')
        break
      case 'bottomRow': numadjmines = bottomRow(todig, sheet, '#fffffe', '#ea9998')
        break
      case 'rightCol': numadjmines = rightCol(todig, sheet, '#fffffe', '#ea9998')
        break
      case 'center': numadjmines = center(todig, sheet, '#fffffe', '#ea9998')
        break
    }
    //Determine how many mines are next to this cell based on its location

    switch(numadjmines) {
      case 1: todig.setValue(1).setFontColor('#1155cc')
        break
      case 2: todig.setValue(2).setFontColor('#38761d')
        break
      case 3: todig.setValue(3).setFontColor('#cc0000')
        break
      case 4: todig.setValue(4).setFontColor('#271cbe')
        break
      case 5: todig.setValue(5).setFontColor('#a61c00')
        break
      case 6: todig.setValue(6).setFontColor('#45818e')
        break
      case 7: todig.setValue(7).setFontColor('#000000')
        break
      case 8: todig.setValue(8).setFontColor('#666666')
        break
      case 9: todig.setValue('what')
        break
      //If this cell has at least one mine next to it, display that number

      case 0: 
        switch(mode) {
          case 'topLeft':
            if (todig.offset(0, 1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
            if (todig.offset(1, 0).getBackground() != '#999999') {recursion(todig.offset(1, 0), sheet, boardheight, boardwidth)}
            if (todig.offset(1, 1).getBackground() != '#999999') {recursion(todig.offset(1, 1), sheet, boardheight, boardwidth)}
            break

          case 'topRight':
            if (todig.offset(0, -1).getBackground() != '#999999') {recursion(todig.offset(0, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(1, -1).getBackground() != '#999999') {recursion(todig.offset(1, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(1, 0).getBackground() != '#999999') {recursion(todig.offset(1, 0), sheet, boardheight, boardwidth)}
            break

          case 'bottomLeft':
            if (todig.offset(-1, 0).getBackground() != '#999999') {recursion(todig.offset(-1, 0), sheet, boardheight, boardwidth)}
            if (todig.offset(-1, 1).getBackground() != '#999999') {recursion(todig.offset(-1, 1), sheet, boardheight, boardwidth)}
            if (todig.offset(0, 1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
            break

          case 'bottomRight':
            if (todig.offset(-1, -1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
            if (todig.offset(-1, 0).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
            if (todig.offset(0, -1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
            break

          case 'topRow':
            if (todig.offset(0, -1).getBackground() != '#999999') {recursion(todig.offset(0, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(0, 1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
            if (todig.offset(1, -1).getBackground() != '#999999') {recursion(todig.offset(1, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(1, 0).getBackground() != '#999999') {recursion(todig.offset(1, 0), sheet, boardheight, boardwidth)}
            if (todig.offset(1, 1).getBackground() != '#999999') {recursion(todig.offset(1, 1), sheet, boardheight, boardwidth)}
            break

          case 'leftCol':
            if (todig.offset(-1, 0).getBackground() != '#999999') {recursion(todig.offset(-1, 0), sheet, boardheight, boardwidth)}
            if (todig.offset(-1, 1).getBackground() != '#999999') {recursion(todig.offset(-1, 1), sheet, boardheight, boardwidth)}
            if (todig.offset(0, 1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
            if (todig.offset(1, 0).getBackground() != '#999999') {recursion(todig.offset(1, 0), sheet, boardheight, boardwidth)}
            if (todig.offset(1, 1).getBackground() != '#999999') {recursion(todig.offset(1, 1), sheet, boardheight, boardwidth)}
            break

          case 'bottomRow':
            if (todig.offset(-1, -1).getBackground() != '#999999') {recursion(todig.offset(-1, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(-1, 0).getBackground() != '#999999') {recursion(todig.offset(-1, 0), sheet, boardheight, boardwidth)}
            if (todig.offset(-1, 1).getBackground() != '#999999') {recursion(todig.offset(-1, 1), sheet, boardheight, boardwidth)}
            if (todig.offset(0, -1).getBackground() != '#999999') {recursion(todig.offset(0, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(0, 1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
            break

          case 'rightCol':
            if (todig.offset(-1, -1).getBackground() != '#999999') {recursion(todig.offset(-1, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(-1, 0).getBackground() != '#999999') {recursion(todig.offset(-1, 0), sheet, boardheight, boardwidth)}
            if (todig.offset(0, -1).getBackground() != '#999999') {recursion(todig.offset(0, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(1, -1).getBackground() != '#999999') {recursion(todig.offset(1, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(1, 0).getBackground() != '#999999') {recursion(todig.offset(1, 0), sheet, boardheight, boardwidth)}
            break

          case 'center':
            if (todig.offset(-1, -1).getBackground() != '#999999') {recursion(todig.offset(-1, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(-1, 0).getBackground() != '#999999') {recursion(todig.offset(-1, 0), sheet, boardheight, boardwidth)}
            if (todig.offset(-1, 1).getBackground() != '#999999') {recursion(todig.offset(-1, 1), sheet, boardheight, boardwidth)}
            if (todig.offset(0, -1).getBackground() != '#999999') {recursion(todig.offset(0, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(0, 1).getBackground() != '#999999') {recursion(todig.offset(0, 1), sheet, boardheight, boardwidth)}
            if (todig.offset(1, -1).getBackground() != '#999999') {recursion(todig.offset(1, -1), sheet, boardheight, boardwidth)}
            if (todig.offset(1, 0).getBackground() != '#999999') {recursion(todig.offset(1, 0), sheet, boardheight, boardwidth)}
            if (todig.offset(1, 1).getBackground() != '#999999') {recursion(todig.offset(1, 1), sheet, boardheight, boardwidth)}
            break

          default: return
        }
    break
    //If this cell has no mines next to it, perform this same function with each neighboring cell that hasn't been dug up
    default: return
    }
  }
}
