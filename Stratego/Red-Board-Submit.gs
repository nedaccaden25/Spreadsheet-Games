function strategosubmitboardred() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Red's Set Up Stratego")
  const redboard = SpreadsheetApp.getActive().getSheetByName("Red's Stratego Board")
  const blueboard = SpreadsheetApp.getActive().getSheetByName("Blue's Stratego Board")

  //Make sure the board to be submitted had been verified and has not already been submitted
  if (sheet.getRange('M9').getValue()!='Board is Valid') {
    Browser.msgBox('Cannot Submit', "You must check your board's validity before you can submit it", Browser.Buttons.OK)
    return
  }

  if (redboard.getRange('M8').getValue()!='Building the Board') {
    Browser.msgBox('Cannot Submit', "You have already submitted your board for this game", Browser.Buttons.OK)
    return
  }

  //Put these values in the red board and color the spaces red on both boards
  redboard.getRange('B7:K9').setValues(sheet.getRange('M6:V8').getValues()).setBackground('#ea9999')
  blueboard.getRange('B7:K9').setBackground('#ea9999')
  Browser.msgBox('Board submitted', `Please go to the "Red's Stratego Board" page now and wait for your opponent to be ready`, Browser.Buttons.OK)

  //Update this player's status
  redboard.getRange('M8').setValue('Ready')
  blueboard.getRange('M8').setValue('Ready')

  //Update the piece counts at the bottom
  redboard.getRange('C13').setValue(1)
  redboard.getRange('D13').setValue(6)
  redboard.getRange('E13').setValue(1)
  redboard.getRange('F13').setValue(2)
  redboard.getRange('G13').setValue(5)
  redboard.getRange('H13').setValue(5)
  redboard.getRange('I13').setValue(2)
  redboard.getRange('J13').setValue(2)
  redboard.getRange('K13').setValue(2)
  redboard.getRange('L13').setValue(1)
  redboard.getRange('M13').setValue(1)
  redboard.getRange('N13').setValue(1)
  redboard.getRange('O13').setValue(1)
  blueboard.getRange('C13').setValue(1)
  blueboard.getRange('D13').setValue(6)
  blueboard.getRange('E13').setValue(1)
  blueboard.getRange('F13').setValue(2)
  blueboard.getRange('G13').setValue(5)
  blueboard.getRange('H13').setValue(5)
  blueboard.getRange('I13').setValue(2)
  blueboard.getRange('J13').setValue(2)
  blueboard.getRange('K13').setValue(2)
  blueboard.getRange('L13').setValue(1)
  blueboard.getRange('M13').setValue(1)
  blueboard.getRange('N13').setValue(1)
  blueboard.getRange('O13').setValue(1)


  //If both players are ready at this point, start the game with red's turn
  if (redboard.getRange('M3').getValue()=='Ready') {
    redboard.getRange('M5').setValue('Your Turn').setBackground('#ea9999')
    blueboard.getRange('M5').setValue("Red's Turn").setBackground('#ea9999')
  }
}
