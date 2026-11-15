function strategomakemoveblue() {
  const blueboard = SpreadsheetApp.getActive().getSheetByName("Blue's Stratego Board")
  const redboard = SpreadsheetApp.getActive().getSheetByName("Red's Stratego Board")

  //Make sure that it's your turn and you have a piece selected
  if (blueboard.getRange('M5').getValue()!='Move it Where?' && blueboard.getRange('M5').getValue()!='Your Turn') {
    Browser.msgBox('It is not your turn', 'You must wait for your opponent to make their move first', Browser.Buttons.OK)
    return
  }
  if (blueboard.getRange('M5').getValue()=='Your Turn') {
    Browser.msgBox('Select a Piece', 'You must select a piece of yours to move', Browser.Buttons.OK)
    return
  }

  var tomove = blueboard.getActiveCell()
  var row = blueboard.getRange('A1').getValue().slice(0,1)
  var column = blueboard.getRange('A1').getValue().slice(2,4)
  var selectedpiece = blueboard.getRange(row, column)
  var piecetype = blueboard.getRange('A2').getValue()
  var enemyvalue = 0
  var whowon = 0
  var counter = 0
  var flag = 0
  var gamewon = 0

  //Handle some more errors
  if (tomove.getBackground()=='#4a86e8') {
    Browser.msgBox('Invalid movement', 'Pieces cannot be moved into the lakes', Browser.Buttons.OK)
    return
  }
  if (tomove.getBackground()=='#a4c2f4') {
    Browser.msgBox('Invalid movement', 'Pieces cannot be moved on top of each other', Browser.Buttons.OK)
    return
  }
  if (tomove.getBackground()=='#ffffff' || tomove.getRow()==1) {
    Browser.msgBox('Invalid movement', 'Pieces must stay within the board', Browser.Buttons.OK)
    return
  }
  if (tomove.getBackground()=='#0000ff') {
    if (!(piecetype==1 && tomove.offset(1,0).getBackground()=='#ea9999')) {
      Browser.msgBox('Invalid movement', 'A piece cannot stay in place as its movement unless it is a Spotter and it is Spotting this turn', Browser.Buttons.OK)
      return
    } else if (piecetype==1 && tomove.offset(1,0).getBackground()=='#ea9999') {flag=1}
  }
  if (tomove.getRow()!=row && tomove.getColumn()!=column) {
    Browser.msgBox('Invalid movement', 'Pieces cannot move diagonally', Browser.Buttons.OK)
    return
  }

  //Differentiate between the two movement types
  if (piecetype!=2 && flag!=1) {
    //Handle one more error
    if (tomove.getRow()-row>1 || tomove.getRow()-row<-1 || tomove.getColumn()-column>1 || tomove.getColumn()-column<-1) {
      Browser.msgBox('Invalid movement', 'This piece cannot move more than one space at a time')
    }

  } else if (piecetype==2) {
    flag=0
    //If the Scout is moving up
    if (tomove.getRow()>row) {
      counter = tomove.getRow()
      while (counter>row) {
        if (!(blueboard.getRange(counter, column).getBackground()=='#b6d7a8' || (blueboard.getRange(counter, column).getBackground()=='#ea9999' && counter==tomove.getRow()))) {
          flag=1
          break
        }
        counter=counter-1
      }
    }
    //If the Scout is moving down
    if (tomove.getRow()<row) {
      counter = tomove.getRow()
      while (counter<row) {
        if (!(blueboard.getRange(counter, column).getBackground()=='#b6d7a8' || (blueboard.getRange(counter, column).getBackground()=='#ea9999' && counter==tomove.getRow()))) {
          flag=1
          break
        }
        counter=counter+1
      }
    }
    //If the Scout is moving right
    if (tomove.getColumn()>column) {
      counter = tomove.getColumn()
      while (counter>column) {
        if (!(blueboard.getRange(row, counter).getBackground()=='#b6d7a8' || (blueboard.getRange(row, counter).getBackground()=='#ea9999' && counter==tomove.getRow()))) {
          flag=1
          break
        }
        counter=counter-1
      }
    }
    //If the Scout is moving left
    if (tomove.getColumn()<column) {
      counter = tomove.getColumn()
      while (counter<column) {
        if (!(blueboard.getRange(row, counter).getBackground()=='#b6d7a8' || (blueboard.getRange(row, counter).getBackground()=='#ea9999' && counter==tomove.getRow()))) {
          flag=1
          break
        }
        counter=counter+1
      }
    }
    if (flag==1) {
      Browser.msgBox('Invalid movement', 'The Scout cannot move through other pieces or across the lakes', Browser.Buttons.OK)
      return
    }
  }
  //If the space is open, move the piece
  if (tomove.getBackground()=='#b6d7a8') {
    tomove.setValue(selectedpiece.getValue()).setBackground('#a4c2f4')
    redboard.getRange(tomove.getRow(),tomove.getColumn()).setBackground('#a4c2f4')
    selectedpiece.setBackground('#b6d7a8').setValue('')
    redboard.getRange(selectedpiece.getRow(),selectedpiece.getColumn()).setBackground('#b6d7a8')
  }

  //If the space is not open, determine who should win this fight
  if (tomove.getBackground()=='#ea9999') {
    enemyvalue = redboard.getRange(tomove.getRow(), tomove.getColumn()).getValue()
    blueboard.getRange('W3').setValue(piecetype)
    blueboard.getRange('T3').setValue(enemyvalue)
    redboard.getRange('W3').setValue(piecetype)
    redboard.getRange('T3').setValue(enemyvalue)

    //If the're both numbers, compare them
    if (Number.isInteger(enemyvalue) && Number.isInteger(piecetype)) {
      if (enemyvalue>piecetype) {
        blueboard.getRange('T4').setValue('Red Won').setBackground('#ea9999')
        redboard.getRange('T4').setValue('Red Won').setBackground('#ea9999')
        whowon='red'
      }
      if (enemyvalue<piecetype) {
        blueboard.getRange('T4').setValue('Blue Won').setBackground('#a4c2f4')
        redboard.getRange('T4').setValue('Blue Won').setBackground('#a4c2f4')
        whowon='blue'
      }
      if (enemyvalue==piecetype) {
        blueboard.getRange('T4').setValue("It's a Tie").setBackground('#ffffff')
        redboard.getRange('T4').setValue("It's a Tie").setBackground('#ffffff')
        whowon='tie'
      }
    }
    //If the current player is using a spy, check for a ten and a spy from the opponent
    else if (piecetype=='S') {
      piecetype=0
      if (enemyvalue==10) {
        blueboard.getRange('T4').setValue('Blue Won').setBackground('#a4c2f4')
        redboard.getRange('T4').setValue('Blue Won').setBackground('#a4c2f4')
        whowon='blue'
      } else if (enemyvalue!='S' && enemyvalue!='s') {
        enemyvalue=0
        blueboard.getRange('T4').setValue('Red Won').setBackground('#ea9999')
        redboard.getRange('T4').setValue('Red Won').setBackground('#ea9999')
        whowon='red'
      } else {
        blueboard.getRange('T4').setValue("It's a Tie").setBackground('#ffffff')
        redboard.getRange('T4').setValue("It's a Tie").setBackground('#ffffff')
        whowon='tie'
      }
    }
    //If you fought a spy with anything other than a spy, you win
    else if (enemyvalue=='S' || enemyvalue=='s') {
      enemyvalue=0
      blueboard.getRange('T4').setValue('Blue Won').setBackground('#a4c2f4')
      redboard.getRange('T4').setValue('Blue Won').setBackground('#a4c2f4')
      whowon='blue'
    }
    //If you fought a bomb with anything other than a three, you lose
    else if (enemyvalue=='B' || enemyvalue=='b') {
      enemyvalue=-1
      if (piecetype==3) {
        blueboard.getRange('T4').setValue('Blue Won').setBackground('#a4c2f4')
        redboard.getRange('T4').setValue('Blue Won').setBackground('#a4c2f4')
        whowon='blue'
      } else {
        blueboard.getRange('T4').setValue('Red Won').setBackground('#ea9999')
        redboard.getRange('T4').setValue('Red Won').setBackground('#ea9999')
        whowon='red'
      }
    }
    //If you fought a flag, you win the entire game
    if (enemyvalue=='F' || enemyvalue=='f') {
      enemyvalue=-2
      blueboard.getRange('T4').setValue('Blue Won').setBackground('#a4c2f4')
      redboard.getRange('T4').setValue('Blue Won').setBackground('#a4c2f4')
      whowon='blue'
      gamewon=1
    }

    //Update the spot with the correct piece
    if (whowon=='blue') {
      blueboard.getRange('E13').offset(0,enemyvalue).setValue(blueboard.getRange('E13').offset(0,enemyvalue).getValue()-1)
      redboard.getRange('E13').offset(0,enemyvalue).setValue(redboard.getRange('E13').offset(0,enemyvalue).getValue()-1)
      tomove.setValue(selectedpiece.getValue()).setBackground('#a4c2f4')
      redboard.getRange(tomove.getRow(),tomove.getColumn()).setBackground('#a4c2f4').setValue('')
      selectedpiece.setBackground('#b6d7a8').setValue('')
      redboard.getRange(selectedpiece.getRow(),selectedpiece.getColumn()).setBackground('#b6d7a8')
    } else if (whowon=='tie') {
      blueboard.getRange('E14').offset(0,piecetype).setValue(blueboard.getRange('E14').offset(0,piecetype).getValue()-1)
      redboard.getRange('E14').offset(0,piecetype).setValue(redboard.getRange('E14').offset(0,piecetype).getValue()-1)
      blueboard.getRange('E13').offset(0,enemyvalue).setValue(blueboard.getRange('E13').offset(0,enemyvalue).getValue()-1)
      redboard.getRange('E13').offset(0,enemyvalue).setValue(redboard.getRange('E13').offset(0,enemyvalue).getValue()-1)
      tomove.setBackground('#b6d7a8')
      redboard.getRange(tomove.getRow(),tomove.getColumn()).setBackground('#b6d7a8').setValue('')
      selectedpiece.setBackground('#b6d7a8').setValue('')
      redboard.getRange(selectedpiece.getRow(),selectedpiece.getColumn()).setBackground('#b6d7a8')
    } else {
      blueboard.getRange('E14').offset(0,piecetype).setValue(blueboard.getRange('E14').offset(0,piecetype).getValue()-1)
      redboard.getRange('E14').offset(0,piecetype).setValue(redboard.getRange('E14').offset(0,piecetype).getValue()-1)
      selectedpiece.setBackground('#b6d7a8').setValue('')
      redboard.getRange(selectedpiece.getRow(),selectedpiece.getColumn()).setBackground('#b6d7a8')
    }
  }
  SpreadsheetApp.flush()
  //Handle the spotter logic
  if (piecetype==1 && tomove.offset(1,0).getBackground()=='#ea9999' && whowon==0) {
    tomove.setBackground('#a4c2f4')
    redboard.getRange(tomove.getRow(),tomove.getColumn()).setBackground('#a4c2f4')
    var msgboxyn = Browser.msgBox('Would you like to spot the piece in front of your spotter?', 'If you say yes, please input your guess into the "Your guess" space and check the checkbox below', Browser.Buttons.YES_NO)

    if (msgboxyn=='yes') {
      var submitbox = blueboard.getRange('X8').isChecked()
      var guesstaken = 0
      var guess = 0
      var guesssuccess = 0
      blueboard.getRange('M5').setValue('Spotting...')
      redboard.getRange('M5').setValue('Spotting...')
      blueboard.getRange('T9').setValue('Spot in Progress').setBackground('#ffffff')
      redboard.getRange('T9').setValue('Spot in Progress').setBackground('#ffffff')

      //Get a valid guess from the player
      while (guesstaken==0) {
        while (!submitbox) {
          SpreadsheetApp.flush()
          submitbox=blueboard.getRange('X8').isChecked()
        }
        submitbox=false
        blueboard.getRange('X8').uncheck()
        guess=blueboard.getRange('X7').getValue()
        if (guess=='f') {guess='F'}
        if (guess=='b') {guess='B'}
        if (guess=='s') {guess='S'}
        if (guess==1 || guess==2 || guess==3 || guess==4 || guess==5 || guess==6 || guess==7 || guess==8 || guess==9 || guess==10 || guess=='S' || guess=='B' || guess=='F') {
          guesstaken=1
          redboard.getRange('X10').setValue(guess)
          tomove.setBackground('#a4c2f4')
          redboard.getRange(tomove.getRow(),tomove.getColumn()).setBackground('#a4c2f4')
        } else {Browser.msgBox('Invalid guess', 'Please input a valid guess', Browser.Buttons.OK)}
      }

      //Check for a successful guess
      if (redboard.getRange(tomove.offset(1,0).getRow(),tomove.offset(1,0).getColumn()).getValue()==guess) {
        guesssuccess=1
      } else if ((!Number.isInteger(guess)) && (!Number.isInteger(redboard.getRange(tomove.offset(1,0).getRow(),tomove.offset(1,0).getColumn()).getValue()))) {
        if ((guess=='F'&&redboard.getRange(tomove.offset(1,0).getRow(),tomove.offset(1,0).getColumn()).getValue()=='f') || (guess=='B'&&redboard.getRange(tomove.offset(1,0).getRow(),tomove.offset(1,0).getColumn()).getValue()=='b') || (guess=='S'&&redboard.getRange(tomove.offset(1,0).getRow(),tomove.offset(1,0).getColumn()).getValue()=='s')) {
          guesssuccess=1
        }
      }

      if (guesssuccess==1) {
        if (guess=='S') {guess=0}
        if (guess=='B') {guess=-1}
        if (guess=='F') {
          guess=-2
          gamewon=1
        }
        tomove.offset(1,0).setBackground('#b6d7a8')
        redboard.getRange(tomove.offset(1,0).getRow(),tomove.offset(1,0).getColumn()).setValue('').setBackground('#b6d7a8')
        blueboard.getRange('E13').offset(0,guess).setValue(blueboard.getRange('E13').offset(0,guess).getValue()-1)
        redboard.getRange('E13').offset(0,guess).setValue(redboard.getRange('E13').offset(0,guess).getValue()-1)
        blueboard.getRange('T9').setValue('Spot Successful').setBackground('#a4c2f4')
        redboard.getRange('T9').setValue('Spot Successful').setBackground('#a4c2f4')
      } else {
        blueboard.getRange('T9').setValue('Spot Failed').setBackground('#ea9999')
        redboard.getRange('T9').setValue('Spot Failed').setBackground('#ea9999')
      }
    //If you didn't want to spot, but your spotter hasn't moved yet, continue your turn instead of passing it
    } else if (flag==1) {
      tomove.setBackground('#a4c2f4')
      redboard.getRange(tomove.getRow(),tomove.getColumn()).setBackground('#a4c2f4')
      Browser.msgBox('Select a different move', 'If your Spotter is not going to Spot this turn, then you must make a move', Browser.Buttons.OK)
      return
    }
  }
  //After everything else has happened, pass the turn or declare a winner
  if (gamewon==0) {
    blueboard.getRange('M5').setValue("Red's Turn").setBackground('#ea9999')
    redboard.getRange('M5').setValue('Your Turn').setBackground('#ea9999')
  } else {
    blueboard.getRange('M5').setValue('Blue Wins!')
    redboard.getRange('M5').setValue('Blue Wins!')
  }
}
