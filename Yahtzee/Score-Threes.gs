function yahtzeethrees() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Yahtzee')

  var currentturn

  var currentcolor = sheet.getRange('A1').getBackground()

  var counter = 0

  var topdone = 1

  var bottomdone = 1

  var tripleyahtzee = 0

  var die1space = sheet.getRange('C3')
  var die2space = sheet.getRange('G3')
  var die3space = sheet.getRange('K3')
  var die4space = sheet.getRange('O3')
  var die5space = sheet.getRange('S3')

  //Concat all the spaces on the die together, this results in a series of dots equal in number to its value

  var die1 = die1space.offset(-1,-1).getValue()+die1space.offset(-1,0).getValue()+die1space.offset(-1,1).getValue()+die1space.offset(0,-1).getValue()+die1space.getValue()+die1space.offset(0,1).getValue()+die1space.offset(1,-1).getValue()+die1space.offset(1,0).getValue()+die1space.offset(1,1).getValue()

  var die2 = die2space.offset(-1,-1).getValue()+die2space.offset(-1,0).getValue()+die2space.offset(-1,1).getValue()+die2space.offset(0,-1).getValue()+die2space.getValue()+die2space.offset(0,1).getValue()+die2space.offset(1,-1).getValue()+die2space.offset(1,0).getValue()+die2space.offset(1,1).getValue()

  var die3 = die3space.offset(-1,-1).getValue()+die3space.offset(-1,0).getValue()+die3space.offset(-1,1).getValue()+die3space.offset(0,-1).getValue()+die3space.getValue()+die3space.offset(0,1).getValue()+die3space.offset(1,-1).getValue()+die3space.offset(1,0).getValue()+die3space.offset(1,1).getValue()

  var die4 = die4space.offset(-1,-1).getValue()+die4space.offset(-1,0).getValue()+die4space.offset(-1,1).getValue()+die4space.offset(0,-1).getValue()+die4space.getValue()+die4space.offset(0,1).getValue()+die4space.offset(1,-1).getValue()+die4space.offset(1,0).getValue()+die4space.offset(1,1).getValue()

  var die5 = die5space.offset(-1,-1).getValue()+die5space.offset(-1,0).getValue()+die5space.offset(-1,1).getValue()+die5space.offset(0,-1).getValue()+die5space.getValue()+die5space.offset(0,1).getValue()+die5space.offset(1,-1).getValue()+die5space.offset(1,0).getValue()+die5space.offset(1,1).getValue()

  //Dertermine who's turn it is, or if we are playing triple yahtzee determine which board to play on
  if (sheet.getRange('Z10').getValue() != 'X1') {
    if (currentcolor == '#ea9999') {currentturn=0}
    else if (currentcolor == '#a4c2f4') {currentturn=1}
    else if (currentcolor == '#b6d7a8') {currentturn=2}
    else if (currentcolor == '#ffe599') {currentturn=3}
    else {return}
  } else {
    if (sheet.getRange('AC3').isChecked()) {counter=counter+1}
    if (sheet.getRange('AC6').isChecked()) {counter=counter+1}
    if (sheet.getRange('AC9').isChecked()) {counter=counter+1}
    if (counter != 1) {
      Browser.msgBox('Please select exactly one board to play on', '', Browser.Buttons.OK)
      return
    }
    counter=0
    tripleyahtzee=1
    if (sheet.getRange('AC3').isChecked()) {currentturn=0}
    if (sheet.getRange('AC6').isChecked()) {currentturn=1}
    if (sheet.getRange('AC9').isChecked()) {currentturn=2}
    sheet.getRange('AC3:AC9').uncheck()
  }
  var currentturndupe = currentturn
  var nextturn = ''
  var nextcolor = ''

  
  if (sheet.getRange('Z3').offset(0, currentturn).getValue() == '' && sheet.getRange('D1').getValue() != 0) {

    //Put the correct value in the box
    if (die1 == '•••') {counter=counter+3}
    if (die2 == '•••') {counter=counter+3}
    if (die3 == '•••') {counter=counter+3}
    if (die4 == '•••') {counter=counter+3}
    if (die5 == '•••') {counter=counter+3}
    if (tripleyahtzee==0) {sheet.getRange('Z3').offset(0, currentturn).setValue(counter)
    } else {sheet.getRange('Z3').offset(0, currentturn).setValue(counter*(currentturn+1))}

    //Handle the Yahtzee bonus section
    if (counter==15 && sheet.getRange('Z16').offset(0, currentturn).getValue()!=0) {
      sheet.getRange('Z18').offset(0, currentturn).setValue(sheet.getRange('Z18').offset(0, currentturn).getValue()+1)
    }

    //Reset roll counter
    sheet.getRange('D1').setValue(0)

    //Uncheck all of the boxes
    sheet.getRange('B6:R6').uncheck()

    //Figure out who's next
    while (nextturn === '' && tripleyahtzee==0) {
      if (currentturn == 3) {currentturn=currentturn-4}
      if (sheet.getRange('Z10').offset(0, currentturn+1).getValue() == 1) {nextturn=currentturn+1}
      else {currentturn=currentturn+1}
    }

    //Color the board corresponding to who's turn it is
    if (nextturn==0 && tripleyahtzee==0) {
      nextcolor='#ea9999'
    }
    if (nextturn==1 && tripleyahtzee==0) {
      nextcolor='#a4c2f4'
    }
    if (nextturn==2 && tripleyahtzee==0) {
      nextcolor='#b6d7a8'
    }
    if (nextturn==3 && tripleyahtzee==0) {
      nextcolor='#ffe599'
    }
    if (tripleyahtzee==0) {
      sheet.getRange('A1:U1').setBackground(nextcolor)
      sheet.getRange('A5:U5').setBackground(nextcolor)
      sheet.getRange('A1:A5').setBackground(nextcolor)
      sheet.getRange('E1:E5').setBackground(nextcolor)
      sheet.getRange('I1:I5').setBackground(nextcolor)
      sheet.getRange('M1:M5').setBackground(nextcolor)
      sheet.getRange('Q1:Q5').setBackground(nextcolor)
      sheet.getRange('U1:U5').setBackground(nextcolor)
      sheet.getRange('B6').setBackground(nextcolor)
      sheet.getRange('F6').setBackground(nextcolor)
      sheet.getRange('J6').setBackground(nextcolor)
      sheet.getRange('N6').setBackground(nextcolor)
      sheet.getRange('R6').setBackground(nextcolor)
    }

    //Figure out if the top of the board is done
    for (i=0;i<6; i++) {
      if (sheet.getRange('Z1').offset(i, currentturndupe).isBlank()) {topdone=0}
    }
    if (topdone==1) {
      //Sum the top part of the board
      counter=0
      counter=counter+sheet.getRange('Z1').offset(0, currentturndupe).getValue()
      counter=counter+sheet.getRange('Z2').offset(0, currentturndupe).getValue()
      counter=counter+sheet.getRange('Z3').offset(0, currentturndupe).getValue()
      counter=counter+sheet.getRange('Z4').offset(0, currentturndupe).getValue()
      counter=counter+sheet.getRange('Z5').offset(0, currentturndupe).getValue()
      counter=counter+sheet.getRange('Z6').offset(0, currentturndupe).getValue()
      sheet.getRange('Z7').offset(0, currentturndupe).setValue(counter)

      //If the requirement was met, apply the bonus score
      if (counter>=63 && tripleyahtzee==0) {
        sheet.getRange('Z8').offset(0, currentturndupe).setValue(35)
        sheet.getRange('Z9').offset(0, currentturndupe).setValue(counter+35)
        sheet.getRange('Z19').offset(0, currentturndupe).setValue(counter+35)
      } else if (counter>=63*(currentturn+1)) {
        sheet.getRange('Z8').offset(0, currentturndupe).setValue(35*(currentturn+1))
        sheet.getRange('Z9').offset(0, currentturndupe).setValue(counter+35*(currentturn+1))
        sheet.getRange('Z19').offset(0, currentturndupe).setValue(counter+35*(currentturn+1))
      }
      else {
        sheet.getRange('Z8').offset(0, currentturndupe).setValue(0)
        sheet.getRange('Z9').offset(0, currentturndupe).setValue(counter)
        sheet.getRange('Z19').offset(0, currentturndupe).setValue(counter)
      }

      //Figure out if the bottom of the board is done
      for (i=0;i<7;i++) {
        if (sheet.getRange('Z11').offset(i, currentturndupe).isBlank()) {bottomdone=0}
      }

      if (bottomdone==1) {
        //Set this player as done if not playing triple yahtzee
        if (tripleyahtzee==0) {
          sheet.getRange('Z10').offset(0, currentturndupe).setValue(0)
        }

        //sum everything on the board
        counter=0
        if (tripleyahtzee==0) {
          counter=counter+sheet.getRange('Z18').offset(0, currentturndupe).getValue()*100
        } else {counter=counter+sheet.getRange('Z18').offset(0, currentturndupe).getValue()*100*(currentturndupe+1)}
        counter=counter+sheet.getRange('Z19').offset(0, currentturndupe).getValue()
        counter=counter+sheet.getRange('Z20').offset(0, currentturndupe).getValue()
        sheet.getRange('Z22').offset(0, currentturndupe).setValue(counter)

        //If this beats the high score, update it for standard yahtzee
        if (counter>sheet.getRange('A22').getValue() && tripleyahtzee==0) {
          sheet.getRange('A22').setValue(counter)

        //Figure out if all boards are done
        } else if (tripleyahtzee==1) {
          if (!sheet.getRange('Z22').isBlank() && !sheet.getRange('AA22').isBlank() && !sheet.getRange('AB22').isBlank()) {

            //Sum every board's total score
            counter=0
            counter=counter+sheet.getRange('Z22').getValue()
            counter=counter+sheet.getRange('AA22').getValue()
            counter=counter+sheet.getRange('AB22').getValue()
            sheet.getRange('AC22').setValue(counter)

            //If this beats the high score, update it for triple yahtzee
            if (counter>sheet.getRange('F22').getValue()) {
              sheet.getRange('F22').setValue(counter)
            }
          }
        }
      }
    }
  }
}
