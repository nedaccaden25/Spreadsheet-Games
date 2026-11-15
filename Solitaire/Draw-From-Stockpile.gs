function solitairedrawfromstock() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Solitaire')

  if (sheet.getRange('A800').getValue()=='1 card at a time') {
    if (sheet.getRange('N2').getBackground()=='#a4c2f4') {
      if (sheet.getRange('O828').getValue()=='♥' || sheet.getRange('O828').getValue()=='♦') {var suitcolor='#ff0000'} else {var suitcolor='#000000'}
      sheet.getRange('L2').setBackground('#ffffff').setValue(sheet.getRange('N828').getValue()).setFontColor(suitcolor)

      rotate(sheet)
      sheet.getRange('R1').setValue(sheet.getRange('R1').getValue()-1)
      if (sheet.getRange('R1').getValue()==0) {
        sheet.getRange('N2').setBackground('#d9ead3')
      }
    } else {
      var num = refresh(sheet)

      sheet.getRange('N2').setBackground('#a4c2f4')
      sheet.getRange('R1').setValue(num)
      sheet.getRange('L2').setBackground('#d9ead3').clearContent()
    }
  } else {
    sheet.getRange('L1:L5').breakApart().setBorder(false, false, false, false, false, false)
    if (sheet.getRange('N2').getBackground()=='#a4c2f4') {
      for (var i=0; i<3; i++) {
        if (sheet.getRange('N2').getBackground()=='#a4c2f4') {
          if (sheet.getRange('O828').getValue()=='♥' || sheet.getRange('O828').getValue()=='♦') {var suitcolor='#ff0000'} else {var suitcolor='#000000'}
          if (i>0) {
            sheet.getRange(i, 12, 3, 1).breakApart()
          }
          sheet.getRange(i+1, 12, 3, 1).merge().setBackground('#ffffff').setValue(sheet.getRange('N828').getValue()).setFontColor(suitcolor).setBorder(true, true, true, true, true, true)

          rotate(sheet)
          sheet.getRange('R1').setValue(sheet.getRange('R1').getValue()-1)
          if (sheet.getRange('R1').getValue()==0) {
            sheet.getRange('N2').setBackground('#d9ead3')
          }
        } else {
          sheet.getRange(i+3, 12).setBackground('#d9ead3').clearContent()
        }
      }
    } else {
      var num = refresh(sheet)

      sheet.getRange('N2').setBackground('#a4c2f4')
      sheet.getRange('R1').setValue(num)
      sheet.getRange('L1:L3').merge().setBackground('#d9ead3').clearContent().setBorder(true, true, true, true, true, true)
      sheet.getRange('L4:L5').setBackground('#d9ead3').clearContent()
    }
  }
  sheet.getRange('W1').setValue(sheet.getRange('W1').getValue()+1)
}

function rotate(sheet) {
  var arr = sheet.getRange('L828:O851').getValues()

  for (var i=23; i>0; i--) {
    arr[i][0]=arr[i-1][0]
  }
  arr[0][0]=arr[0][2]

  for (var i=23; i>0; i--) {
    arr[i][1]=arr[i-1][1]
  }
  arr[0][1]=arr[0][3]

  for (var i=0; i<23; i++) {
    arr[i][2]=arr[i+1][2]
  }
  arr[23][2]=''

  for (var i=0; i<23; i++) {
    arr[i][3]=arr[i+1][3]
  }
  arr[23][3]=''

  sheet.getRange('L828:O851').setValues(arr).setFontColor('#ffffff')
}

function refresh(sheet) {
  for (var i=0; i>-1; i++) {
    if (!sheet.getRange('L851').offset(-i, 0).isBlank()) {
      var temp=i
      break
    }
  }

  var arr = []
  while (temp<24) {
    arr.push(sheet.getRange('L851').offset(-temp, 0, 1, 2).getValues()[0])
    temp++
  }
  var num = arr.length
  for (i=24-num; i!=0; i--) {
    arr.push(['',''])
  }
Logger.log(arr)
  sheet.getRange('N828:O851').setValues(arr).setFontColor('#ffffff')
  sheet.getRange('L828:M851').setValue('').setFontColor('#ffffff')

  return(num)
}
