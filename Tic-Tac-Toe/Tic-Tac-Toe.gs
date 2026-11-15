function tictactoenewgame() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Tic-Tac-Toe')
  
  var board = [[], [], []]
  var boardcolors = [['#ffffff', '#ffffff', '#ffffff'], ['#ffffff', '#ffffff', '#ffffff'], ['#ffffff', '#ffffff', '#ffffff']]
  var checks
  var checkspot
  var turn
  var color
  var winner
  var flag = 0
  var movecount = 0

  //Get the ranges that the board consits of in an array to reference later
  for (i=0; i<3; i++) {
    for (j=0; j<3; j++) {
      board[i].push(sheet.getRange('B2').offset(i, j))
    }
  }

  //Initialize board
  sheet.getRange('A1:E5').setBackground('#ffffff')
  sheet.getRange('B2:D4').setBorder(true, true, true, true, true, true).setFontSize(23).insertCheckboxes().uncheck()

  //Initialize turn display
  sheet.getRange('F2:I2').merge().setBackground('#ea9999').setValue("Red's Turn").setFontSize(27)

  while (flag==0) {
    SpreadsheetApp.flush()
    checks = sheet.getRange('B2:D4').getValues()
    for (i=0; i<3; i++) {
      for (j=0; j<3; j++) {
        Logger.log(checks[i][j])
        if (checks[i][j]==true) {
          board[i][j].uncheck()
          if (boardcolors[i][j]=='#ffffff') {
            movecount++
            checkspot = sheet.getRange('B2').offset(i, j)
            color = sheet.getRange('F2').getBackground()
            if (color=='#ea9999') {turn='Red'} else {turn='Blue'}
            checkspot.setBackground(color)
            boardcolors[i][j]=color
            winner = checkforwin(sheet, turn, board, checks, checkspot, color, boardcolors, i, j)
            if (winner==0 && movecount!=9) {
              if (turn=='Red') {
                sheet.getRange('F2').setValue("Blue's Turn").setBackground('#a4c2f4')
              } else {
                
                sheet.getRange('F2').setValue("Red's Turn").setBackground('#ea9999')
              }
            } else {
              flag=1
            }
          }
        }
        if (flag==1) {break}
      }
      if (flag==1) {break}
    }
  }
  if (winner==0) {
    sheet.getRange('F2').setValue('The game is a CAT').setBackground('#ffffff')
  } else {
    sheet.getRange('F2').setValue(turn+' is the Winner!')
  }
}

function checkforwin(turn, color, boardcolors, i, j) {
  var up = checkdirection(color, boardcolors, i, j, 1, 0)
  var left = checkdirection(color, boardcolors, i, j, 0, -1)
  var down = checkdirection(color, boardcolors, i, j, -1, 0)
  var right = checkdirection(color, boardcolors, i, j, 0, 1)
  var upleft = checkdirection(color, boardcolors, i, j, 1, -1)
  var downleft = checkdirection(color, boardcolors, i, j, -1, -1)
  var downright = checkdirection(color, boardcolors, i, j, -1, 1)
  var upright = checkdirection(color, boardcolors, i, j, 1, 1)

  if (up+down==2) {return(turn)}
  if (left+right==2) {return(turn)}
  if (upleft+downright==2) {return(turn)}
  if (downleft+upright==2) {return(turn)}
  return(0)
}

function checkdirection(color, boardcolors, i, j, vert, horz) {
  if (i-vert>-1 && i-vert<3 && j+horz>-1 && j+horz<3) {
    if (boardcolors[i-vert][j+horz]==color) {return(1+checkdirection(color, boardcolors, i-vert, j+horz, vert, horz))}
  }
  return(0)
}
