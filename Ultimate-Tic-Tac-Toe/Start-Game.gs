function ultimatetictactoestart() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Ultimate Tic-Tac-Toe')

  //Initialize board
  sheet.getRange('B2:J10').insertCheckboxes().uncheck().setBackground('#ffffff')
  sheet.getRange('L2').setValue("Red's Turn").setBackground('#ea9999')
  sheet.getRange('S3:U5').setBackground('#ffffff')

  //Initialize variables
  var currentturn = '#ea9999'
  var boardswon = [['#ffffff', '#ffffff', '#ffffff'], ['#ffffff', '#ffffff', '#ffffff'], ['#ffffff', '#ffffff', '#ffffff']]
  var boardcolors = []
  var boardcolors2d = []
  var checkpos = [0, 0]
  var smallwin = 0
  var bigwin = 0

  //Build the 3D array of board colors
  for (i=0; i<9; i++) {
    boardcolors.push([])
    for (j=0; j<3; j++) {
      boardcolors[i].push([])
      for (k=0; k<3; k++) {
        boardcolors[i][j].push('#ffffff')
      }
    }
  }

  while(true) {

    //Wait for a check, then get its position
    checkpos = waitforcheck(sheet)
    checkpos = convertto3d(checkpos)

    //Update the array with the proper color
    boardcolors[checkpos[0]][checkpos[1]][checkpos[2]] = currentturn

    //Check if this 3x3 has been won
    smallwin = checkwin(boardcolors[checkpos[0]], currentturn)

    if (smallwin!=0) {
      //Update the corresponding arrays with this win
      if (smallwin==1) {
        boardswon[Math.floor(checkpos[0]/3)][checkpos[0]%3] = currentturn
        boardcolors[checkpos[0]] = colorafterwin(boardcolors[checkpos[0]], currentturn)
      } else {
        boardswon[Math.floor(checkpos[0]/3)][checkpos[0]%3] = '#cccccc'
        boardcolors[checkpos[0]] = colorafterwin(boardcolors[checkpos[0]], '#cccccc')
      }

      //Check if the whole game has been won
      bigwin = checkwin(boardswon, currentturn)

      smallwin=0
    }

    //Update the sheet's appearance
    boardcolors2d = convertarrto2d(boardcolors)
    sheet.getRange('B2:J10').setBackgrounds(boardcolors2d)
    boardcolors2d = []

    //If a winner was determined, break out of the main while loop
    if (bigwin!=0) {
      break
    }

    //Figure out which board the next player should play in
    if (boardswon[checkpos[1]][checkpos[2]]=='#ffffff') {
      sheet.getRange('S3:U5').setBackground('#000000')
      sheet.getRange('S3').offset(checkpos[1], checkpos[2]).setBackground('#ffffff')
    } else {
      sheet.getRange('S3:U5').setBackground('#ffffff')
    }

    //Pass the turn
    currentturn = passturn(sheet, currentturn)
  }

  //Display the winning or CAT text
  if (bigwin==1) {
    if (currentturn == '#ea9999') {
      sheet.getRange('L2').setValue('Red Wins')
    } else {
      sheet.getRange('L2').setValue('Blue Wins')
    }
  } else {
    sheet.getRange('L2').setValue('CAT').setBackground('#cccccc')
  }
}

function waitforcheck(sheet) {
  while (true) {
    SpreadsheetApp.flush()
    var checks = sheet.getRange('B2:J10').getValues()
    for (i=0; i<9; i++) {
      for (j=0; j<9; j++) {
        if (checks[i][j]==true) {
          if (sheet.getRange('B2').offset(i, j).getBackground()=='#ffffff' && sheet.getRange('S3').offset(Math.floor(i/3), Math.floor(j/3)).getBackground()=='#ffffff') {
            sheet.getRange('B2').offset(i, j).uncheck()
            return([i, j])
          } else {
            sheet.getRange('B2').offset(i, j).uncheck()
            Browser.msgBox('Invalid Move', '', Browser.Buttons.OK)
          }
        }
      }
    }
  }
}

function convertto3d(toconvert) {
  var j = toconvert[0]%3
  var k = toconvert[1]%3
  var i = Math.floor(toconvert[0]/3)*3+Math.floor(toconvert[1]/3)
  return([i, j, k])
}

function checkwin(board, checkagainst) {
  var arr = []
  var empty = 0
  var toreturn

  for (i=0; i<3; i++) {
    arr.push([])
    for (j=0; j<3; j++) {
      if (board[i][j]==checkagainst) {arr[i].push(1)} else {arr[i].push(0)}
      if (board[i][j]=='#ffffff') {empty=1}
    }
  }

  toreturn = threeinarow(arr)

  if (toreturn==0 && empty==0) {
    return(2)
  } else {
    return(toreturn)
  }
}

function threeinarow (arr) {
  if (arr[0][0]==1 && arr[0][1]==1 && arr[0][2]==1) {return(1)}
  if (arr[1][0]==1 && arr[1][1]==1 && arr[1][2]==1) {return(1)}
  if (arr[2][0]==1 && arr[2][1]==1 && arr[2][2]==1) {return(1)}
  if (arr[0][0]==1 && arr[1][0]==1 && arr[2][0]==1) {return(1)}
  if (arr[0][1]==1 && arr[1][1]==1 && arr[2][1]==1) {return(1)}
  if (arr[0][2]==1 && arr[1][2]==1 && arr[2][2]==1) {return(1)}
  if (arr[0][0]==1 && arr[1][1]==1 && arr[2][2]==1) {return(1)}
  if (arr[0][2]==1 && arr[1][1]==1 && arr[2][0]==1) {return(1)}
  return(0)
}

function colorafterwin (smallboardcolors, color) {
  for (i=0; i<3; i++) {
    for (j=0; j<3; j++) {
      if (color == '#ea9999') {
        if (smallboardcolors[i][j] == '#ffffff') {smallboardcolors[i][j]='#f4cccc'}
        else if (smallboardcolors[i][j] == '#ea9999') {smallboardcolors[i][j]='#ff0000'}
        else if (smallboardcolors[i][j] == '#a4c2f4') {smallboardcolors[i][j]='#8e7cc3'}
      } else if (color == '#a4c2f4') {
        if (smallboardcolors[i][j] == '#ffffff') {smallboardcolors[i][j]='#c9daf8'}
        else if (smallboardcolors[i][j] == '#ea9999') {smallboardcolors[i][j]='#8e7cc3'}
        else if (smallboardcolors[i][j] == '#a4c2f4') {smallboardcolors[i][j]='#0000ff'}
      } else {
        if (smallboardcolors[i][j] == '#ffffff') {smallboardcolors[i][j]='#999999'}
        else if (smallboardcolors[i][j] == '#ea9999') {smallboardcolors[i][j]='#995555'}
        else if (smallboardcolors[i][j] == '#a4c2f4') {smallboardcolors[i][j]='#555599'}
      }
    }
  }
  return(smallboardcolors)
}

function passturn (sheet, currentturn) {
  if (currentturn=='#ea9999') {
    sheet.getRange('L2').setValue("Blue's Turn").setBackground('#a4c2f4')
    return('#a4c2f4')
  } else {
    sheet.getRange('L2').setValue("Red's Turn").setBackground('#ea9999')
    return('#ea9999')
  }
}

function convertarrto2d (boardcolors) {
  var arr = []
  for (i=0; i<9; i++) {
    arr.push([])
    for (j=0; j<9; j++) {
      arr[i].push(boardcolors[Math.floor(i/3)*3+Math.floor(j/3)][i%3][j%3])
    }
  }
  return(arr)
}
