function strategoboardvalidblue() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Blue's Set Up Stratego")

  var placedpieces = sheet.getRange('M1:V3').getValues()

  var num10 = 0
  var num9 = 0
  var num8 = 0
  var num7 = 0
  var num6 = 0
  var num5 = 0
  var num4 = 0
  var num3 = 0
  var num2 = 0
  var num1 = 0
  var numS = 0
  var numB = 0
  var numF = 0
  var numinvalid = 0
  var toomanylist = 'There are too many '
  var notenoughlist = 'There are not enough '

  //Count up the currently placed pieces
  for (i=0; i<3; i++) {
    for (j=0; j<10; j++) {
      switch (placedpieces[i][j]) {
        case '10': num10=num10+1
          sheet.getRange('M1').offset(i,j).setValue('="10"')
          break
        case 10: num10=num10+1
          sheet.getRange('M1').offset(i,j).setValue('="10"')
          break
        case '9': num9=num9+1
          sheet.getRange('M1').offset(i,j).setValue('="9"')
          break
        case 9: num9=num9+1
          sheet.getRange('M1').offset(i,j).setValue('="9"')
          break
        case '8': num8=num8+1
          sheet.getRange('M1').offset(i,j).setValue('="8"')
          break
        case 8: num8=num8+1
          sheet.getRange('M1').offset(i,j).setValue('="8"')
          break
        case '7': num7=num7+1
          sheet.getRange('M1').offset(i,j).setValue('="7"')
          break
        case 7: num7=num7+1
          sheet.getRange('M1').offset(i,j).setValue('="7"')
          break
        case '6': num6=num6+1
          sheet.getRange('M1').offset(i,j).setValue('="6"')
          break
        case 6: num6=num6+1
          sheet.getRange('M1').offset(i,j).setValue('="6"')
          break
        case '5': num5=num5+1
          sheet.getRange('M1').offset(i,j).setValue('="5"')
          break
        case 5: num5=num5+1
          sheet.getRange('M1').offset(i,j).setValue('="5"')
          break
        case '4': num4=num4+1
          sheet.getRange('M1').offset(i,j).setValue('="4"')
          break
        case 4: num4=num4+1
          sheet.getRange('M1').offset(i,j).setValue('="4"')
          break
        case '3': num3=num3+1
          sheet.getRange('M1').offset(i,j).setValue('="3"')
          break
        case 3: num3=num3+1
          sheet.getRange('M1').offset(i,j).setValue('="3"')
          break
        case '2': num2=num2+1
          sheet.getRange('M1').offset(i,j).setValue('="2"')
          break
        case 2: num2=num2+1
          sheet.getRange('M1').offset(i,j).setValue('="2"')
          break
        case '1': num1=num1+1
          sheet.getRange('M1').offset(i,j).setValue('="1"')
          break
        case 1: num1=num1+1
          sheet.getRange('M1').offset(i,j).setValue('="1"')
          break
        case 'S': numS=numS+1
        sheet.getRange('M1').offset(i,j).setValue('="S"')
          break
        case 's': numS=numS+1
        sheet.getRange('M1').offset(i,j).setValue('="s"')
          break
        case 'B': numB=numB+1
        sheet.getRange('M1').offset(i,j).setValue('="B"')
          break
        case 'b': numB=numB+1
        sheet.getRange('M1').offset(i,j).setValue('="b"')
          break
        case 'F': numF=numF+1
        sheet.getRange('M1').offset(i,j).setValue('="F"')
          break
        case 'f': numF=numF+1
        sheet.getRange('M1').offset(i,j).setValue('="f"')
          break
        default: numinvalid=numinvalid+1
          sheet.getRange('M1').offset(i,j).setValue('')
      }
    }
  }

  //Check if everything is correct
  if (num10==1 && num9==1 && num8==1 && num7==1 && num6==2 && num5==2 && num4==2 && num3==5 && num2==5 && numS==1 && numB==6 && numF==1) {
    Browser.msgBox('Your board is valid', '', Browser.Buttons.OK)
    sheet.getRange('M9').setValue('=IF(AND(M1="'+placedpieces[0][0]+'", N1="'+placedpieces[0][1]+'", O1="'+placedpieces[0][2]+'", P1="'+placedpieces[0][3]+'", Q1="'+placedpieces[0][4]+'", R1="'+placedpieces[0][5]+'", S1="'+placedpieces[0][6]+'", T1="'+placedpieces[0][7]+'", U1="'+placedpieces[0][8]+'", V1="'+placedpieces[0][9]+'", M2="'+placedpieces[1][0]+'", N2="'+placedpieces[1][1]+'", O2="'+placedpieces[1][2]+'", P2="'+placedpieces[1][3]+'", Q2="'+placedpieces[1][4]+'", R2="'+placedpieces[1][5]+'", S2="'+placedpieces[1][6]+'", T2="'+placedpieces[1][7]+'", U2="'+placedpieces[1][8]+'", V2="'+placedpieces[1][9]+'", M3="'+placedpieces[2][0]+'", N3="'+placedpieces[2][1]+'", O3="'+placedpieces[2][2]+'", P3="'+placedpieces[2][3]+'", Q3="'+placedpieces[2][4]+'", R3="'+placedpieces[2][5]+'", S3="'+placedpieces[2][6]+'", T3="'+placedpieces[2][7]+'", U3="'+placedpieces[2][8]+'", V3="'+placedpieces[2][9]+'"), "Board is Valid", "Please Check Board Validity")')
    return
  } else {sheet.getRange('M9').setValue('Please Check Board Validity')}

  //Determine if there are too many or not enough of the certain pieces
  if (num10>1) {toomanylist=toomanylist+'tens, '}
  else if (num10<1) {notenoughlist=notenoughlist+'tens, '}

  if (num9>1) {toomanylist=toomanylist+'nines, '}
  else if (num9<1) {notenoughlist=notenoughlist+'nines, '}

  if (num8>1) {toomanylist=toomanylist+'eights, '}
  else if (num8<1) {notenoughlist=notenoughlist+'eights, '}
  
  if (num7>1) {toomanylist=toomanylist+'sevens, '}
  else if (num7<1) {notenoughlist=notenoughlist+'sevens, '}
  
  if (num6>2) {toomanylist=toomanylist+'sixes, '}
  else if (num6<2) {notenoughlist=notenoughlist+'sixes, '}
  
  if (num5>2) {toomanylist=toomanylist+'fives, '}
  else if (num5<2) {notenoughlist=notenoughlist+'fives, '}
  
  if (num4>2) {toomanylist=toomanylist+'fours, '}
  else if (num4<2) {notenoughlist=notenoughlist+'fours, '}
  
  if (num3>5) {toomanylist=toomanylist+'threes, '}
  else if (num3<5) {notenoughlist=notenoughlist+'threes, '}
  
  if (num2>5) {toomanylist=toomanylist+'twos, '}
  else if (num2<5) {notenoughlist=notenoughlist+'twos, '}
  
  if (num1>2) {toomanylist=toomanylist+'ones, '}
  else if (num1<2) {notenoughlist=notenoughlist+'ones, '}
  
  if (numS>1) {toomanylist=toomanylist+'spies, '}
  else if (numS<1) {notenoughlist=notenoughlist+'spies, '}
  
  if (numB>6) {toomanylist=toomanylist+'bombs, '}
  else if (numB<6) {notenoughlist=notenoughlist+'bombs, '}
  
  if (numF>1) {toomanylist=toomanylist+'flags, '}
  else if (numF<1) {notenoughlist=notenoughlist+'flags, '}

  //Remove the leading comma from the lists
  notenoughlist=notenoughlist.slice(0,-2)
  toomanylist=toomanylist.slice(0,-2)

  //Display any error messages that need to be shown
  if (toomanylist!='There are too man') {
    Browser.msgBox('There are too many of some pieces', toomanylist, Browser.Buttons.OK)
  }

  if (notenoughlist!='There are not enoug') {
    Browser.msgBox('There are not enough of some pieces', notenoughlist, Browser.Buttons.OK)
  }

  if (numinvalid!=0) {
    Browser.msgBox('Some of your spaces are blank or are not valid pieces', numinvalid+' spaces are blank or invalid, these spaces have been emptied', Browser.Buttons.OK)
  }
}
