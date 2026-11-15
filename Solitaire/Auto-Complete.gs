function solitaireautocomplete() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Solitaire')

  var row7colors = sheet.getRange('B7:N7').getBackgrounds()[0]
  var incomplete = 0

  if (sheet.getRange('B7').getValue()=='CONGRATULATIONS!') {return}

  for (i=0; i<14; i=i+2) {
    if (row7colors[i]=='#a4c2f4') {incomplete=1}
  }

  if (!sheet.getRange('L1').isBlank() || !sheet.getRange('L2').isBlank() || !sheet.getRange('L3').isBlank() || sheet.getRange('N2').getBackground()!='#d9ead3') {incomplete=1}

  if (incomplete==1) {
    Browser.msgBox('The board is not yet complete.', 'All cards must be flipped over and the stockpile must be empty to auto-complete.', Browser.Buttons.OK)
    return
  } else {
    var heartnum = sheet.getRange('B2').getValue()
    var diamondnum = sheet.getRange('D2').getValue()
    var spadenum = sheet.getRange('F2').getValue()
    var clubnum = sheet.getRange('H2').getValue()

    heartnum = heartnum.slice(0, heartnum.length-1)
    diamondnum = diamondnum.slice(0, diamondnum.length-1)
    spadenum = spadenum.slice(0, spadenum.length-1)
    clubnum = clubnum.slice(0, clubnum.length-1)

    heartnum = numberize(heartnum)
    diamondnum = numberize(diamondnum)
    spadenum = numberize(spadenum)
    clubnum = numberize(clubnum)

    var movestoadd = (13-heartnum)+(13-diamondnum)+(13-spadenum)+(13-clubnum)

    sheet.getRange('W1').setValue(sheet.getRange('W1').getValue()+movestoadd)
    sheet.getRange('B7:N26').breakApart().clearContent().setBackground('#d9ead3').setFontColor('#000000').setBorder(false, false, false, false, false, false)
    sheet.getRange('B7:N9').merge().setValue('CONGRATULATIONS!').setFontSize(37).setHorizontalAlignment('center').setBorder(true, true, true, true, true, true)
    sheet.getRange('B2').setValue('K♥').setFontColor('#ff0000').setBackground('#ffffff')
    sheet.getRange('D2').setValue('K♦').setFontColor('#ff0000').setBackground('#ffffff')
    sheet.getRange('F2').setValue('K♠').setFontColor('#000000').setBackground('#ffffff')
    sheet.getRange('H2').setValue('K♣').setFontColor('#000000').setBackground('#ffffff')
    if (sheet.getRange('W1').getValue()<sheet.getRange('A29').getValue()) {
      sheet.getRange('A29').setValue(sheet.getRange('W1').getValue())
    }
  }
}

function numberize (num) {
  if (num=='A') {return(1)}
  if (num=='J') {return(11)}
  if (num=='Q') {return(12)}
  if (num=='K') {return(13)}
  return(num)
}
