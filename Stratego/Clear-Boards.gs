function strategoclearboard() {
  const redboard = SpreadsheetApp.getActive().getSheetByName("Red's Stratego Board")
  const blueboard = SpreadsheetApp.getActive().getSheetByName("Blue's Stratego Board")

  var yesno = Browser.msgBox('Are you sure you want to clear the boards?', 'This will completely reset both of the boards', Browser.Buttons.YES_NO)

  if (yesno=='yes') {
    //Clear the red board
    redboard.getRange('B2:K9').setValue('').setBackground('#b6d7a8').setFontColor('#000000')
    redboard.getRange('D5:E6').setValue('').setBackground('#4a86e8')
    redboard.getRange('H5:I6').setValue('').setBackground('#4a86e8')
    redboard.getRange('C13:O14').setValue('')
    redboard.getRange('T3').setValue('')
    redboard.getRange('W3').setValue('')
    redboard.getRange('T4').setValue('').setBackground('#ffffff')
    redboard.getRange('X7').setValue('')
    redboard.getRange('X8').uncheck()
    redboard.getRange('T9').setValue('').setBackground('#ffffff')
    redboard.getRange('X10').setValue('')
    redboard.getRange('M3').setValue('Building the Board')
    redboard.getRange('M8').setValue('Building the Board')
    redboard.getRange('M5').setValue('Waiting for Boards').setBackground('#ffffff')

    //Clear the blue board
    blueboard.getRange('B2:K9').setValue('').setBackground('#b6d7a8').setFontColor('#000000')
    blueboard.getRange('D5:E6').setValue('').setBackground('#4a86e8')
    blueboard.getRange('H5:I6').setValue('').setBackground('#4a86e8')
    blueboard.getRange('C13:O14').setValue('')
    blueboard.getRange('T3').setValue('')
    blueboard.getRange('W3').setValue('')
    blueboard.getRange('T4').setValue('').setBackground('#ffffff')
    blueboard.getRange('X7').setValue('')
    blueboard.getRange('X8').uncheck()
    blueboard.getRange('T9').setValue('').setBackground('#ffffff')
    blueboard.getRange('X10').setValue('')
    blueboard.getRange('M3').setValue('Building the Board')
    blueboard.getRange('M8').setValue('Building the Board')
    blueboard.getRange('M5').setValue('Waiting for Boards').setBackground('#ffffff')
  }
}
