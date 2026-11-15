function strategosubmitboardblue() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Blue's Set Up Stratego")
  const blueboard = SpreadsheetApp.getActive().getSheetByName("Blue's Stratego Board")
  const redboard = SpreadsheetApp.getActive().getSheetByName("Red's Stratego Board")

  //Make sure the board to be submitted had been verified and has not already been submitted
  if (sheet.getRange('M9').getValue()!='Board is Valid') {
    Browser.msgBox('Cannot Submit', "You must check your board's validity before you can submit it", Browser.Buttons.OK)
    return
  }

  if (blueboard.getRange('M3').getValue()!='Building the Board') {
    Browser.msgBox('Cannot Submit', "You have already submitted your board for this game", Browser.Buttons.OK)
    return
  }

  //Put these values in the blue board and color the spaces blue on both boards
  blueboard.getRange('B2:K4').setValues(sheet.getRange('M1:V3').getValues()).setBackground('#a4c2f4')
  redboard.getRange('B2:K4').setBackground('#a4c2f4')
  Browser.msgBox('Board submitted', `Please go to the "Blue's Stratego Board" page now and wait for your opponent to be ready`, Browser.Buttons.OK)

  //Update this player's status
  blueboard.getRange('M3').setValue('Ready')
  redboard.getRange('M3').setValue('Ready')

  //Update the piece counts at the bottom
  blueboard.getRange('C14').setValue(1)
  blueboard.getRange('D14').setValue(6)
  blueboard.getRange('E14').setValue(1)
  blueboard.getRange('F14').setValue(2)
  blueboard.getRange('G14').setValue(5)
  blueboard.getRange('H14').setValue(5)
  blueboard.getRange('I14').setValue(2)
  blueboard.getRange('J14').setValue(2)
  blueboard.getRange('K14').setValue(2)
  blueboard.getRange('L14').setValue(1)
  blueboard.getRange('M14').setValue(1)
  blueboard.getRange('N14').setValue(1)
  blueboard.getRange('O14').setValue(1)
  redboard.getRange('C14').setValue(1)
  redboard.getRange('D14').setValue(6)
  redboard.getRange('E14').setValue(1)
  redboard.getRange('F14').setValue(2)
  redboard.getRange('G14').setValue(5)
  redboard.getRange('H14').setValue(5)
  redboard.getRange('I14').setValue(2)
  redboard.getRange('J14').setValue(2)
  redboard.getRange('K14').setValue(2)
  redboard.getRange('L14').setValue(1)
  redboard.getRange('M14').setValue(1)
  redboard.getRange('N14').setValue(1)
  redboard.getRange('O14').setValue(1)


  //If both players are ready at this point, start the game with red's turn
  if (blueboard.getRange('M8').getValue()=='Ready') {
    blueboard.getRange('M5').setValue("Red's Turn").setBackground('#ea9999')
    redboard.getRange('M5').setValue('Your Turn').setBackground('#ea9999')
  }
}
