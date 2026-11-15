function solitairemakemove() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Solitaire')

  var boardcolors = sheet.getRange('A1:O27').getBackgrounds()
  var selectedcells = []
  var selectedcards = []
  var tomove = sheet.getActiveCell()
  var tomovecard = tomove.getValue()
  var tomovecolor = tomove.getBackground()
  var tomovecol = tomove.getColumn()
  var tomoverow = tomove.getRow()
  var movetype = ''

  if (tomovecolor == '#ea9999') {
    for (i=0; i<27; i++) {
      for (j=0; j<15; j++) {
        if (boardcolors[i][j]=='#ea9999') {
          sheet.getRange('A1').offset(i, j).setBackground('#ffffff')
        }
      }
    }
    return
  }

  //Handle errors
  for (i=0; i<27; i++) {
    for (j=0; j<15; j++) {
      if (boardcolors[i][j]=='#ea9999') {
        selectedcells.push(sheet.getRange('A1').offset(i, j))
        selectedcards.push(sheet.getRange('A1').offset(i, j).getValue())
      }
    }
  }

  if (selectedcells.length==0) {
    Browser.msgBox('Please select a card to move', '', Browser.Buttons.OK)
    return
  }

  if ((tomovecolor=='#d9ead3' && !(tomoverow==2 && (tomovecol==2 || tomovecol==4 || tomovecol==6 || tomovecol==8)) && !(tomoverow==7 && (tomovecol==2 || tomovecol==4 || tomovecol==6 || tomovecol==8 || tomovecol==10 || tomovecol==12 || tomovecol==14))) || tomove.getRow()>27 || tomove.getColumn()>15 || (tomove.getRow()<6 && tomove.getColumn()==12) ||tomovecolor=='#a4c2f4') {
    Browser.msgBox('Please select a valid place to move the card(s)', '', Browser.Buttons.OK)
    return
  }

  if (tomoverow==2 && selectedcards.length>1) {
    Browser.msgBox('Only one card can be moved into the foundation at a time', '', Browser.Buttons.OK)
    return
  }

  //Isolate the number and suit to make things easier
  var tomovenumber = tomovecard.slice(0, tomovecard.length-1)
  if (tomovenumber=='A') {tomovenumber=1}
  if (tomovenumber=='J') {tomovenumber=11}
  if (tomovenumber=='Q') {tomovenumber=12}
  if (tomovenumber=='K') {tomovenumber=13}
  var tomovesuit = tomovecard.slice(tomovecard.length-1, tomovecard.length)
  var selectednumber = selectedcards[0].slice(0,selectedcards[0].length-1)
  if (selectednumber=='A') {selectednumber=1}
  if (selectednumber=='J') {selectednumber=11}
  if (selectednumber=='Q') {selectednumber=12}
  if (selectednumber=='K') {selectednumber=13}
  var selectedsuit = selectedcards[0].slice(selectedcards[0].length-1,selectedcards[0].length)
  var failed = 0

  //Check if this is a valid move

  if (tomoverow!=2) {
    movetype='tableau'

    if (tomovenumber-selectednumber!=1 && tomovenumber-selectednumber!=-13) {
      failed=1
    }

    if (!((tomovesuit=='♥' && (selectedsuit=='♠' || selectedsuit=='♣')) || (tomovesuit=='♦' && (selectedsuit=='♠' || selectedsuit=='♣')) || (tomovesuit=='♠' && (selectedsuit=='♥' || selectedsuit=='♦')) || (tomovesuit=='♣' && (selectedsuit=='♥' || selectedsuit=='♦'))) && selectednumber!=13) {
      failed=1
    }
  } else {
    tomovesuit=tomove.offset(-1, 0).getValue()
    movetype='foundation'

    if (tomovesuit!=selectedsuit) {
      failed=1
    }

    if (selectednumber-tomovenumber!=1) {
      failed=1
    }
  }

  if (failed==1) {
    Browser.msgBox('This is not a valid move', 'Make sure to double check the suits and numbers of these cards', Browser.Buttons.OK)
    return
  }

  //Make the move depending on which type of move you're doing
  if (movetype=='foundation') {
    if (tomovesuit=='♥' || tomovesuit=='♦') {var suitcolor='#ff0000'} else {var suitcolor='#000000'}
    tomove.setValue(selectedcards[0]).setBackground('#ffffff').setFontColor(suitcolor)
  } else {
    if (tomovecolor!='#d9ead3') {
      tomove.offset(0, 0, 3, 1).breakApart().setBorder(null, null, false, null, null, null)
      if (selectedsuit=='♥' || selectedsuit=='♦') {var suitcolor='#ff0000'} else {var suitcolor='#000000'}
      tomove.offset(1, 0, 3, 1).merge().setValue(selectedcards[0]).setBackground('#ffffff').setFontColor(suitcolor).setBorder(true, true, true, true, true, true)
      var i = 1
    } else {
      if (selectedsuit=='♥' || selectedsuit=='♦') {var suitcolor='#ff0000'} else {var suitcolor='#000000'}
      tomove.setValue(selectedcards[0]).setBackground('#ffffff').setFontColor(suitcolor)
      var i = 0
    }
    while (selectedcards.length>1) {
      if (selectedcards[1].slice(selectedcards[1].length-1, selectedcards[1].length)=='♥' || selectedcards[1].slice(selectedcards[1].length-1, selectedcards[1].length)=='♦') {suitcolor='#ff0000'} else {suitcolor='#000000'}
      tomove.offset(i, 0, 3, 1).breakApart().setBorder(null, null, false, null, null, null)
      i=i+1
      tomove.offset(i, 0, 3, 1).merge().setValue(selectedcards[1]).setBackground('#ffffff').setFontColor(suitcolor).setBorder(true, true, true, true, true, true)
      selectedcards.splice(1, 1)
    }
  }

  //Remove the selected cards from their initial spots
  while (selectedcells.length>1) {
    selectedcells[selectedcells.length-1].setBackground('#d9ead3').clearContent().offset(0, 0, 3, 1).breakApart().setBorder(null, false, false, false, null, null)
    selectedcells.splice(selectedcells.length-1, 1)
    selectedcells[selectedcells.length-1].offset(0, 0, 3, 1).merge().setBorder(true, true, true, true, true, true)
  }

  if (selectedcells[0].offset(-1, 0).getBackground()!='#d9ead3' && selectedcells[0].getRow()>6) {
    selectedcells[0].setBackground('#d9ead3').clearContent().offset(0, 0, 3, 1).breakApart().setBorder(null, false, false, false, null, null)
    selectedcells[0].offset(-1, 0, 3, 1).merge().setBorder(true, true, true, true, true, true)
    if (selectedcells[0].offset(-1, 0).getBackground()=='#a4c2f4') {
      var temp = selectedcells[0].offset(799, 0).getValue()
      if (temp.slice(temp.length-1, temp.length)=='♥' || temp.slice(temp.length-1, temp.length)=='♦') {suitcolor='#ff0000'} else {suitcolor='#000000'}
      selectedcells[0].offset(-1, 0).setBackground('#ffffff').setValue(selectedcells[0].offset(799, 0).getValue()).setFontColor(suitcolor)
    }
  } else {
    if (selectedcells[0].getRow()<6 && selectedcells[0].getColumn()==12) {
      if (sheet.getRange('A800').getValue()=='1 card at a time') {
        var arr = sheet.getRange('L828:M851').getValues()
        arr.splice(0, 1)
        sheet.getRange('L828:M850').setValues(arr).setFontColor('#ffffff')
        sheet.getRange('L851:M851').clearContent()
        if (arr[0][0]=='') {
          selectedcells[0].setBackground('#d9ead3').clearContent()
        } else {
          if (arr[0][1]=='♥' || arr[0][1]=='♦') {suitcolor='#ff0000'} else {suitcolor='#000000'}
          selectedcells[0].setValue(arr[0][0]).setFontColor(suitcolor).setBackground('#ffffff')

        }
      } else {
        var arr = sheet.getRange('L828:M851').getValues()
        arr.splice(0, 1)
        sheet.getRange('L828:M850').setValues(arr).setFontColor('#ffffff')
        sheet.getRange('L851:M851').clearContent()

        if (selectedcells[0].getRow()==3) {
          if (arr[0][1]=='♥' || arr[0][1]=='♦') {suitcolor='#ff0000'} else {suitcolor='#000000'}
          selectedcells[0].setValue(arr[0][0]).setFontColor(suitcolor).setBackground('#ffffff')
          if (selectedcells[0].isBlank()) {selectedcells[0].setBackground('#d9ead3').setBorder(false, false, false, false, null, null)}
          if (arr[1][1]=='♥' || arr[1][1]=='♦') {suitcolor='#ff0000'} else {suitcolor='#000000'}
          selectedcells[0].offset(-1, 0).setValue(arr[1][0]).setFontColor(suitcolor)
          if (selectedcells[0].offset(-1, 0).isBlank()) {selectedcells[0].offset(-1, 0).setBackground('#d9ead3').setBorder(false, false, null, false, null, null)}
          if (arr[2][1]=='♥' || arr[2][1]=='♦') {suitcolor='#ff0000'} else {suitcolor='#000000'}
          selectedcells[0].offset(-2, 0).setValue(arr[2][0]).setFontColor(suitcolor)
          if (selectedcells[0].offset(-2, 0).isBlank()) {selectedcells[0].offset(-2, 0).setBackground('#d9ead3').setBorder(false, false, null, false, null, null)}

        } else if (selectedcells[0].getRow()==2) {
          if (arr[0][1]=='♥' || arr[0][1]=='♦') {suitcolor='#ff0000'} else {suitcolor='#000000'}
          selectedcells[0].setValue(arr[0][0]).setFontColor(suitcolor).setBackground('#ffffff')
          if (selectedcells[0].isBlank()) {selectedcells[0].setBackground('#d9ead3').setBorder(false, false, false, false, null, null)}
          if (arr[1][1]=='♥' || arr[1][1]=='♦') {suitcolor='#ff0000'} else {suitcolor='#000000'}
          selectedcells[0].offset(-1, 0).setValue(arr[1][0]).setFontColor(suitcolor)
          if (selectedcells[0].offset(-1, 0).isBlank()) {selectedcells[0].offset(-1, 0).setBackground('#d9ead3').setBorder(false, false, null, false, null, null)}

        } else {
          if (arr[0][1]=='♥' || arr[0][1]=='♦') {suitcolor='#ff0000'} else {suitcolor='#000000'}
          selectedcells[0].setValue(arr[0][0]).setFontColor(suitcolor).setBackground('#ffffff')
          if (selectedcells[0].isBlank()) {selectedcells[0].setBackground('#d9ead3').setBorder(false, false, false, false, null, null)}
        }
      }
    } else if (selectedcells[0].getRow()==2) {
      if (selectednumber=='1') {
        selectedcells[0].setBackground('#d9ead3').clearContent()
      } else {
        if (selectednumber==11) {selectednumber='J'}
        if (selectednumber==12) {selectednumber='Q'}
        if (selectednumber==13) {selectednumber='K'}
        selectedcells[0].setValue((selectednumber-1)+selectedsuit).setBackground('#ffffff')
      }
    } else {
      selectedcells[0].setBackground('#d9ead3').clearContent()
    }
  }


  //Increment the move counter
  sheet.getRange('W1').setValue(sheet.getRange('W1').getValue()+1)

  //Check for a win
  if (sheet.getRange('B2').getValue()=='K♥' && sheet.getRange('D2').getValue()=='K♦' && sheet.getRange('F2').getValue()=='K♠' && sheet.getRange('H2').getValue()=='K♣') {
    sheet.getRange('B7:N9').merge().setValue('CONGRATULATIONS!').setFontColor('#000000').setBackground('#d9ead3').setFontSize(37).setHorizontalAlignment('center')
    if (sheet.getRange('W1').getValue()<sheet.getRange('A29').getValue()) {
      sheet.getRange('A29').setValue(sheet.getRange('W1').getValue())
    }
  }
}
