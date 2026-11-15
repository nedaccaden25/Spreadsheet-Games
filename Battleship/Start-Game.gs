function battleshipstartgame() {
  const sheet = [SpreadsheetApp.getActive().getSheetByName('Battleship Player 1'), SpreadsheetApp.getActive().getSheetByName('Battleship Player 2')]

  if (sheet[0].getRange('O12').getBackground()!='#b6d7a8') {
    Browser.msgBox('Cannot start game yet', 'You must first make sure your board is valid by clicking the Check Board button', Browser.Buttons.OK)
    return
  }

  if (sheet[0].getRange('B12').getBackground()!='#b6d7a8') {
    Browser.msgBox('Cannot start game yet', 'You must wait for your opponent to make sure their board is valid', Browser.Buttons.OK)
    return
  }

  sheet[0].getRange('B2:K11').uncheck().setBackground('#ffffff')
  sheet[1].getRange('B2:K11').uncheck().setBackground('#ffffff')

  var currentturn = 1
  var opponentturn = 0
  var winflag = 0
  var mode = sheet[0].getRange('J19').getValue()
  var boathit = ''
  var temp = ''
  var sunkflag = 0
  var goagain = 0
  
  //Get the positions of all of the boats
  boats = []

  //Initialize P1 variables
  var p1board = sheet[0].getRange('O2:X11').getValues()
  var p1carrier = []
  var p1battleship = []
  var p1cruiser = []
  var p1submarine = []
  var p1destroyer = []
  var p1boats = []

  //Store the locations of each of the ships for P1
  for (i=0; i<10; i++) {
    for (j=0; j<10; j++) {
      if (p1board[i][j].toUpperCase()=='C') {p1carrier.push([i, j])}
      if (p1board[i][j].toUpperCase()=='B') {p1battleship.push([i, j])}
      if (p1board[i][j].toUpperCase()=='R') {p1cruiser.push([i, j])}
      if (p1board[i][j].toUpperCase()=='S') {p1submarine.push([i, j])}
      if (p1board[i][j].toUpperCase()=='D') {p1destroyer.push([i, j])}
    }
  }
  p1boats = [p1carrier, p1battleship, p1cruiser, p1submarine, p1destroyer]

  //Initialize P2 variables
  var p2board = sheet[1].getRange('O2:X11').getValues()
  var p2carrier = []
  var p2battleship = []
  var p2cruiser = []
  var p2submarine = []
  var p2destroyer = []
  var p2boats = []

  //Store the locations of each of the ships for P1
  for (i=0; i<10; i++) {
    for (j=0; j<10; j++) {
      if (p2board[i][j].toUpperCase()=='C') {p2carrier.push([i, j])}
      if (p2board[i][j].toUpperCase()=='B') {p2battleship.push([i, j])}
      if (p2board[i][j].toUpperCase()=='R') {p2cruiser.push([i, j])}
      if (p2board[i][j].toUpperCase()=='S') {p2submarine.push([i, j])}
      if (p2board[i][j].toUpperCase()=='D') {p2destroyer.push([i, j])}
    }
  }
  p2boats = [p2carrier, p2battleship, p2cruiser, p2submarine, p2destroyer]

  //Consolidate all of the boat positions to be referenced later
  boats = [p1boats, p2boats]
  Logger.log(boats)

  while (true) {
    //Set the turn display to be correct
    if (goagain==0) {
      temp = swapturn(sheet, currentturn, opponentturn)
      currentturn = temp[0]
      opponentturn = temp[1]
      temp=''
    }
    goagain=0

    //Wait until a box has been checked and get the position of that check
    position = waitforcheck(sheet[currentturn])

    //Check which boat was hit, if any
    var i = 0
    for (i=0; i<5; i++) {
      temp = checkforhit(sheet, currentturn, opponentturn, boats[opponentturn][i], position, i)
      boats[opponentturn][i] = temp[0]
      boathit = temp[1]
      temp=''
      if (boathit!=='') {break}
    }

    //If a boat was hit, check if it was sunk by that hit and check if the game should be over
    if (boathit!=='') {

      if (mode == 'After a HIT, play again') {
        goagain=1
      }

      for (i=0; i<boats[opponentturn][boathit].length; i++) {
        if (boats[opponentturn][boathit][i][0] != 'hit') {
          sunkflag=1
          break
        }
      }

      for (i=0; i<5; i++) {
        for (j=0; j<boats[opponentturn][i].length; j++) {
          if (boats[opponentturn][i][j][0] != 'hit') {
            winflag=1
            break
          }
        }
        if (winflag==1) {break}
      }

      if (sunkflag==0) {
        sheet[currentturn].getRange('L13').setValue('SUNK!').setBackground('#cc0000')
        sheet[opponentturn].getRange('L13').setValue('SUNK!').setBackground('#cc0000')
      } else {
        sheet[currentturn].getRange('L13').setValue('HIT!').setBackground('#ea9999')
        sheet[opponentturn].getRange('L13').setValue('HIT!').setBackground('#ea9999')
      }

      if (winflag==0) {
        break
      }

      //Reset flags for the next turn
      winflag = 0
      sunkflag = 0
    }

    //If no boat was hit, color the spaces and label the miss
    if (boathit==='') {
      sheet[currentturn].getRange('B2').offset(position[0], position[1]).setBackground('#444444')
      sheet[opponentturn].getRange('O2').offset(position[0], position[1]).setBackground('#ea9999')
      sheet[currentturn].getRange('L13').setValue('miss').setBackground('#ffffff')
      sheet[opponentturn].getRange('L13').setValue('miss').setBackground('#ffffff')
    }

    //Reset the boathit variable for the next turn
    boathit = ''
  }
  sheet[currentturn].getRange('J14').setValue('You Win!').setFontSize(40)
  sheet[opponentturn].getRange('J14').setValue('You Lose').setFontSize(35)
  sheet[0].getRange('B12:O12').setBackground('#ffffff')
  sheet[1].getRange('B12:O12').setBackground('#ffffff')
  sheet[0].getRange('L13').clearContent().setBackground('#ffffff')
}

function swapturn(sheet, currentturn, opponentturn) {
  sheet[currentturn].getRange('J14').setValue("Opponent's Turn").setFontSize(22)
  sheet[opponentturn].getRange('J14').setValue('Your Turn').setFontSize(32)
  return([opponentturn, currentturn])
}

function waitforcheck(sheet) {
  while (true) {
    SpreadsheetApp.flush()
    var checks = sheet.getRange('B2:K11').getValues()
    for (i=0; i<10; i++) {
      for (j=0; j<10; j++) {
        if (checks[i][j]==true) {
          if (sheet.getRange('B2').offset(i, j).getBackground()=='#ffffff') {
            sheet.getRange('B2').offset(i, j).uncheck()
            return([i, j])
          } else {
            sheet.getRange('B2').offset(i, j).uncheck()
          }
        }
      }
    }
  }
}

function checkforhit(sheet, currentturn, opponentturn, boat, position, tocheck) {
  //Check the position of the check against each of the opponent's boat spots
  for (i=0; i<boat.length; i++) {
    if (position[0]==boat[i][0] && position[1]==boat[i][1]){
      sheet[opponentturn].getRange('O2').offset(position[0], position[1]).setBackground('#ff0000')
      sheet[currentturn].getRange('B2').offset(position[0], position[1]).setBackground('#ff0000')
      boat[i]=['hit', '']
      return([boat, tocheck])
    }
  }
  return([boat, ''])
}
