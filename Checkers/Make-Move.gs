function checkersmakemove() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Checkers')
  var counter = 0
  var locationx = 0
  var locationy = 0
  var newlocation = 0
  var numblues = 0
  var numreds = 0

  //Get the color and box status of every relevant cell on the board

  var a1col = sheet.getRange('A1').getBackground()
  var a1box = sheet.getRange('A1').isChecked()
  var c1col = sheet.getRange('C1').getBackground()
  var c1box = sheet.getRange('C1').isChecked()
  var e1col = sheet.getRange('E1').getBackground()
  var e1box = sheet.getRange('E1').isChecked()
  var g1col = sheet.getRange('G1').getBackground()
  var g1box = sheet.getRange('G1').isChecked()
  var b2col = sheet.getRange('B2').getBackground()
  var b2box = sheet.getRange('B2').isChecked()
  var d2col = sheet.getRange('D2').getBackground()
  var d2box = sheet.getRange('D2').isChecked()
  var f2col = sheet.getRange('F2').getBackground()
  var f2box = sheet.getRange('F2').isChecked()
  var h2col = sheet.getRange('H2').getBackground()
  var h2box = sheet.getRange('H2').isChecked()
  var a3col = sheet.getRange('A3').getBackground()
  var a3box = sheet.getRange('A3').isChecked()
  var c3col = sheet.getRange('C3').getBackground()
  var c3box = sheet.getRange('C3').isChecked()
  var e3col = sheet.getRange('E3').getBackground()
  var e3box = sheet.getRange('E3').isChecked()
  var g3col = sheet.getRange('G3').getBackground()
  var g3box = sheet.getRange('G3').isChecked()
  var b4col = sheet.getRange('B4').getBackground()
  var b4box = sheet.getRange('B4').isChecked()
  var d4col = sheet.getRange('D4').getBackground()
  var d4box = sheet.getRange('D4').isChecked()
  var f4col = sheet.getRange('F4').getBackground()
  var f4box = sheet.getRange('F4').isChecked()
  var h4col = sheet.getRange('H4').getBackground()
  var h4box = sheet.getRange('H4').isChecked()
  var a5col = sheet.getRange('A5').getBackground()
  var a5box = sheet.getRange('A5').isChecked()
  var c5col = sheet.getRange('C5').getBackground()
  var c5box = sheet.getRange('C5').isChecked()
  var e5col = sheet.getRange('E5').getBackground()
  var e5box = sheet.getRange('E5').isChecked()
  var g5col = sheet.getRange('G5').getBackground()
  var g5box = sheet.getRange('G5').isChecked()
  var b6col = sheet.getRange('B6').getBackground()
  var b6box = sheet.getRange('B6').isChecked()
  var d6col = sheet.getRange('D6').getBackground()
  var d6box = sheet.getRange('D6').isChecked()
  var f6col = sheet.getRange('F6').getBackground()
  var f6box = sheet.getRange('F6').isChecked()
  var h6col = sheet.getRange('H6').getBackground()
  var h6box = sheet.getRange('H6').isChecked()
  var a7col = sheet.getRange('A7').getBackground()
  var a7box = sheet.getRange('A7').isChecked()
  var c7col = sheet.getRange('C7').getBackground()
  var c7box = sheet.getRange('C7').isChecked()
  var e7col = sheet.getRange('E7').getBackground()
  var e7box = sheet.getRange('E7').isChecked()
  var g7col = sheet.getRange('G7').getBackground()
  var g7box = sheet.getRange('G7').isChecked()
  var b8col = sheet.getRange('B8').getBackground()
  var b8box = sheet.getRange('B8').isChecked()
  var d8col = sheet.getRange('D8').getBackground()
  var d8box = sheet.getRange('D8').isChecked()
  var f8col = sheet.getRange('F8').getBackground()
  var f8box = sheet.getRange('F8').isChecked()
  var h8col = sheet.getRange('H8').getBackground()
  var h8box = sheet.getRange('H8').isChecked()

  //Put these into arrays to make them easier to access

  var colorlist = new Array(8)
  var boxeslist = new Array(8)

  colorlist = [[a1col, '#cccccc', c1col, '#cccccc', e1col, '#cccccc', g1col, '#cccccc'], ['#cccccc', b2col, '#cccccc', d2col, '#cccccc', f2col, '#cccccc', h2col], [a3col, '#cccccc', c3col, '#cccccc', e3col, '#cccccc', g3col, '#cccccc'], ['#cccccc', b4col, '#cccccc', d4col, '#cccccc', f4col, '#cccccc', h4col], [a5col, '#cccccc', c5col, '#cccccc', e5col, '#cccccc', g5col, '#cccccc'], ['#cccccc', b6col, '#cccccc', d6col, '#cccccc', f6col, '#cccccc', h6col], [a7col, '#cccccc', c7col, '#cccccc', e7col, '#cccccc', g7col, '#cccccc'], ['#cccccc', b8col, '#cccccc', d8col, '#cccccc', f8col, '#cccccc', h8col]]

  boxeslist = [[a1box, false, c1box, false, e1box, false, g1box, false], [false, b2box, false, d2box, false, f2box, false, h2box], [a3box, false, c3box, false, e3box, false, g3box, false], [false, b4box, false, d4box, false, f4box, false, h4box], [a5box, false, c5box, false, e5box, false, g5box, false], [false, b6box, false, d6box, false, f6box, false, h6box], [a7box, false, c7box, false, e7box, false, g7box, false], [false, b8box, false, d8box, false, f8box, false, h8box]]

  //Who's turn is it?
  if (sheet.getRange('I1').getValue() == "Red's turn") {
    var turn = 'red'
  } else if (sheet.getRange('I1').getValue() == "Blue's turn"){
    var turn = 'blue'
  } else {return}

  //What are the relevant spaces?
  //Y=0 IS THE TOP, X=0 IS THE LEFT
  for (i=0; i<8; i++) {
    for (j=0; j<8; j++) {
      if (boxeslist[i][j]==true && (colorlist[i][j]=='#ea9999' || colorlist[i][j]=='#ff0000') && turn=='red') {
        counter=counter+1
        locationx=j
        locationy=i
      }
      if (boxeslist[i][j]==true && (colorlist[i][j]=='#a4c2f4' || colorlist[i][j]=='#4a86e8') && turn=='blue') {
        counter=counter+1
        locationx=j
        locationy=i
      }
    }
  }

  //Handle errors
  if (counter==0) {
    Browser.msgBox('Please select a piece of your color to move', '', Browser.Buttons.OK)
    return
  }
  if (counter>1) {
    Browser.msgBox('Please only select one of your pieces', '', Browser.Buttons.OK)
    return
  }

  //Move the piece and store it's new location
  newlocation = domove(sheet, turn, locationy, locationx, boxeslist, colorlist)

  //If a move happened, uncheck the boxes, otherwise display an error
  if (newlocation.getRow()==locationy+1 && newlocation.getColumn()==locationx+1) {
    Browser.msgBox('Please select a valid move for that piece', 'If no move is available, you must forfeit the game', Browser.Buttons.OK)
    return
  } else {
    sheet.getRange('A1:H8').uncheck()
  }

  if (newlocation.getRow()==1 && turn=='red') {
    newlocation.setBackground('#ff0000')
  }
  if (newlocation.getRow()==8 && turn=='blue') {
    newlocation.setBackground('#4a86e8')
  }

  //Detect if there are no more pieces of a certain color
  for (i=1; i<=8; i++) {
    for (j=1; j<=8; j++) {
      if (sheet.getRange(i, j).getBackground()=='#a4c2f4' || sheet.getRange(i, j).getBackground()=='#4a86e8') {
        numblues=numblues+1
      }
      if (sheet.getRange(i, j).getBackground()=='#ea9999' || sheet.getRange(i, j).getBackground()=='#ff0000') {
        numreds=numreds+1
      }
    }
  }

  //Display a win message if there are none of a certain color
  if (numblues == 0) {
    sheet.getRange('I1').setValue('Red Wins!')
    return
  }
  if (numreds == 0) {
    sheet.getRange('I1').setValue('Blue Wins!')
    return
  }
Logger.log(turn)
  //If the game isn't over, pass the turn
  if (turn=='red') {
      sheet.getRange('I1').setValue("Blue's turn").setBackground('#a4c2f4')
    } else {
    sheet.getRange('I1').setValue("Red's turn").setBackground('#ea9999')
    }

}

function domove(sheet, turn, locationy, locationx, boxeslist, colorlist) {

  var flag = 0

  //Check for a standard move
  if (flag==0 && locationy>0 && locationx<7) {
    if (turn=='red' && boxeslist[locationy-1][locationx+1]==true && colorlist[locationy-1][locationx+1]=='#ffffff') {
      flag=1
      sheet.getRange(locationy, locationx+2).setBackground(sheet.getRange(locationy+1, locationx+1).getBackground())
      sheet.getRange(locationy+1, locationx+1).setBackground('#ffffff')
      return(sheet.getRange(locationy, locationx+2))
    }
  }
  if (flag==0 && locationy>0 && locationx>0) {
    if (turn=='red' && boxeslist[locationy-1][locationx+-1]==true && colorlist[locationy-1][locationx-1]=='#ffffff') {
      flag=1
      sheet.getRange(locationy, locationx).setBackground(sheet.getRange(locationy+1, locationx+1).getBackground())
      sheet.getRange(locationy+1, locationx+1).setBackground('#ffffff')
      return(sheet.getRange(locationy, locationx))
    }
  }
  if (flag==0 && locationy<7 && locationx<7) {
    if (turn=='blue' && boxeslist[locationy+1][locationx+1]==true && colorlist[locationy+1][locationx+1]=='#ffffff') {
      flag=1
      sheet.getRange(locationy+2, locationx+2).setBackground(sheet.getRange(locationy+1, locationx+1).getBackground())
      sheet.getRange(locationy+1, locationx+1).setBackground('#ffffff')
      return(sheet.getRange(locationy+2, locationx+2))
    }
  }
  if (flag==0 && locationy<7 && locationx>0) {
    if (turn=='blue' && boxeslist[locationy+1][locationx-1]==true && colorlist[locationy+1][locationx-1]=='#ffffff') {
      flag=1
      sheet.getRange(locationy+2, locationx).setBackground(sheet.getRange(locationy+1, locationx+1).getBackground())
      sheet.getRange(locationy+1, locationx+1).setBackground('#ffffff')
      return(sheet.getRange(locationy+2, locationx))
    }
  }

  //Check for a king moving backwards
  if (flag==0 && locationy>0 && locationx<7) {
    if (turn=='blue' && colorlist[locationy][locationx]=='#4a86e8' && boxeslist[locationy-1][locationx+1]==true && colorlist[locationy-1][locationx+1]=='#ffffff') {
      flag=1
      sheet.getRange(locationy, locationx+2).setBackground(sheet.getRange(locationy+1, locationx+1).getBackground())
      sheet.getRange(locationy+1, locationx+1).setBackground('#ffffff')
      return(sheet.getRange(locationy, locationx+2))
    }
  }
  if (flag==0 && locationy>0 && locationx>0) {
    if (turn=='blue' && colorlist[locationy][locationx]=='#4a86e8' && boxeslist[locationy-1][locationx+-1]==true && colorlist[locationy-1][locationx-1]=='#ffffff') {
      flag=1
      sheet.getRange(locationy, locationx).setBackground(sheet.getRange(locationy+1, locationx+1).getBackground())
      sheet.getRange(locationy+1, locationx+1).setBackground('#ffffff')
      return(sheet.getRange(locationy, locationx))
    }
  }
  if (flag==0 && locationy<7 && locationx<7) {
    if (turn=='red' && colorlist[locationy][locationx]=='#ff0000' && boxeslist[locationy+1][locationx+1]==true && colorlist[locationy+1][locationx+1]=='#ffffff') {
      flag=1
      sheet.getRange(locationy+2, locationx+2).setBackground(sheet.getRange(locationy+1, locationx+1).getBackground())
      sheet.getRange(locationy+1, locationx+1).setBackground('#ffffff')
      return(sheet.getRange(locationy+2, locationx+2))
    }
  }
  if (flag==0 && locationy<7 && locationx>0) {
    if (turn=='red' && colorlist[locationy][locationx]=='#ff0000' && boxeslist[locationy+1][locationx-1]==true && colorlist[locationy+1][locationx-1]=='#ffffff') {
      flag=1
      sheet.getRange(locationy+2, locationx).setBackground(sheet.getRange(locationy+1, locationx+1).getBackground())
      sheet.getRange(locationy+1, locationx+1).setBackground('#ffffff')
      return(sheet.getRange(locationy+2, locationx))
    }
  }

  //Check for a forward jump
  if (flag==0 && locationy>1 && locationx<6) {
    if (turn=='red' && boxeslist[locationy-2][locationx+2]==true && colorlist[locationy-2][locationx+2]=='#ffffff' && (colorlist[locationy-1][locationx+1]=='#a4c2f4' || colorlist[locationy-1][locationx+1]=='#4a86e8')) {
      flag=1
      sheet.getRange(locationy-1, locationx+3).setBackground(sheet.getRange(locationy+1, locationx+1).getBackground())
      colorlist[locationy-2][locationx+2]=sheet.getRange(locationy+1, locationx+1).getBackground()
      sheet.getRange(locationy, locationx+2).setBackground('#ffffff')
      colorlist[locationy-1][locationx+1]='#ffffff'
      sheet.getRange(locationy+1, locationx+1).setBackground('#ffffff')
      colorlist[locationy][locationx]='#ffffff'
      return(domove(sheet, turn, locationy-2, locationx+2, boxeslist, colorlist))
    }
  }
  if (flag==0 && locationy>1 && locationx>1) {
    if (turn=='red' && boxeslist[locationy-2][locationx-2]==true && colorlist[locationy-2][locationx-2]=='#ffffff' && (colorlist[locationy-1][locationx-1]=='#a4c2f4' || colorlist[locationy-1][locationx-1]=='#4a86e8')) {
      flag=1
      sheet.getRange(locationy-1, locationx-1).setBackground(sheet.getRange(locationy+1, locationx+1).getBackground())
      colorlist[locationy-2][locationx-2]=sheet.getRange(locationy+1, locationx+1).getBackground()
      sheet.getRange(locationy, locationx).setBackground('#ffffff')
      colorlist[locationy-1][locationx-1]='#ffffff'
      sheet.getRange(locationy+1, locationx+1).setBackground('#ffffff')
      colorlist[locationy][locationx]='#ffffff'
      return(domove(sheet, turn, locationy-2, locationx-2, boxeslist, colorlist))
    }
  }
  if (flag==0 && locationy<6 && locationx<6) {
    if (turn=='blue' && boxeslist[locationy+2][locationx+2]==true && colorlist[locationy+2][locationx+2]=='#ffffff' && (colorlist[locationy+1][locationx+1]=='#ea9999' || colorlist[locationy+1][locationx+1]=='#ff0000')) {
      flag=1
      sheet.getRange(locationy+3, locationx+3).setBackground(sheet.getRange(locationy+1, locationx+1).getBackground())
      colorlist[locationy+2][locationx+2]=sheet.getRange(locationy+1, locationx+1).getBackground()
      sheet.getRange(locationy+2, locationx+2).setBackground('#ffffff')
      colorlist[locationy+1][locationx+1]='#ffffff'
      sheet.getRange(locationy+1, locationx+1).setBackground('#ffffff')
      colorlist[locationy][locationx]='#ffffff'
      return(domove(sheet, turn, locationy+2, locationx+2, boxeslist, colorlist))
    }
  }
  if (flag==0 && locationy<6 && locationx>1) {
    if (turn=='blue' && boxeslist[locationy+2][locationx-2]==true && colorlist[locationy+2][locationx-2]=='#ffffff' && (colorlist[locationy+1][locationx-1]=='#ea9999' || colorlist[locationy+1][locationx-1]=='#ff0000')) {
      flag=1
      sheet.getRange(locationy+3, locationx-1).setBackground(sheet.getRange(locationy+1, locationx+1).getBackground())
      colorlist[locationy+2][locationx-2]=sheet.getRange(locationy+1, locationx+1).getBackground()
      sheet.getRange(locationy+2, locationx).setBackground('#ffffff')
      colorlist[locationy+1][locationx-1]='#ffffff'
      sheet.getRange(locationy+1, locationx+1).setBackground('#ffffff')
      colorlist[locationy][locationx]='#ffffff'
      return(domove(sheet, turn, locationy+2, locationx-2, boxeslist, colorlist))
    }
  }

  //Check for a backward jump
  if (flag==0 && locationy>1 && locationx<6) {
    if (turn=='blue' && colorlist[locationy][locationx]=='#4a86e8' && boxeslist[locationy-2][locationx+2]==true && colorlist[locationy-2][locationx+2]=='#ffffff' && (colorlist[locationy-1][locationx+1]=='#ea9999' || colorlist[locationy-1][locationx+1]=='#ff0000')) {
      flag=1
      sheet.getRange(locationy-1, locationx+3).setBackground(sheet.getRange(locationy+1, locationx+1).getBackground())
      colorlist[locationy-2][locationx+2]=sheet.getRange(locationy+1, locationx+1).getBackground()
      sheet.getRange(locationy, locationx+2).setBackground('#ffffff')
      colorlist[locationy-1][locationx+1]='#ffffff'
      sheet.getRange(locationy+1, locationx+1).setBackground('#ffffff')
      colorlist[locationy][locationx]='#ffffff'
      return(domove(sheet, turn, locationy-2, locationx+2, boxeslist, colorlist))
    }
  }
  if (flag==0 && locationy>1 && locationx>1) {
    if (turn=='blue' && colorlist[locationy][locationx]=='#4a86e8' && boxeslist[locationy-2][locationx-2]==true && colorlist[locationy-2][locationx-2]=='#ffffff' && (colorlist[locationy-1][locationx-1]=='#ea9999' || colorlist[locationy-1][locationx-1]=='#ff0000')) {
      flag=1
      sheet.getRange(locationy-1, locationx-1).setBackground(sheet.getRange(locationy+1, locationx+1).getBackground())
      colorlist[locationy-2][locationx-2]=sheet.getRange(locationy+1, locationx+1).getBackground()
      sheet.getRange(locationy, locationx).setBackground('#ffffff')
      colorlist[locationy-1][locationx-1]='#ffffff'
      sheet.getRange(locationy+1, locationx+1).setBackground('#ffffff')
      colorlist[locationy][locationx]='#ffffff'
      return(domove(sheet, turn, locationy-2, locationx-2, boxeslist, colorlist))
    }
  }
  if (flag==0 && locationy<6 && locationx<6) {
    if (turn=='red' && colorlist[locationy][locationx]=='#ff0000' && boxeslist[locationy+2][locationx+2]==true && colorlist[locationy+2][locationx+2]=='#ffffff' && (colorlist[locationy+1][locationx+1]=='#a4c2f4' || colorlist[locationy+1][locationx+1]=='#4a86e8')) {
      flag=1
      sheet.getRange(locationy+3, locationx+3).setBackground(sheet.getRange(locationy+1, locationx+1).getBackground())
      colorlist[locationy+2][locationx+2]=sheet.getRange(locationy+1, locationx+1).getBackground()
      sheet.getRange(locationy+2, locationx+2).setBackground('#ffffff')
      colorlist[locationy+1][locationx+1]='#ffffff'
      sheet.getRange(locationy+1, locationx+1).setBackground('#ffffff')
      colorlist[locationy][locationx]='#ffffff'
      return(domove(sheet, turn, locationy+2, locationx+2, boxeslist, colorlist))
    }
  }
  if (flag==0 && locationy<6 && locationx>1) {
    if (turn=='red' && colorlist[locationy][locationx]=='#ff0000' && boxeslist[locationy+2][locationx-2]==true && colorlist[locationy+2][locationx-2]=='#ffffff' && (colorlist[locationy+1][locationx-1]=='#a4c2f4' || colorlist[locationy+1][locationx-1]=='#4a86e8')) {
      flag=1
      sheet.getRange(locationy+3, locationx-1).setBackground(sheet.getRange(locationy+1, locationx+1).getBackground())
      colorlist[locationy+2][locationx-2]=sheet.getRange(locationy+1, locationx+1).getBackground()
      sheet.getRange(locationy+2, locationx).setBackground('#ffffff')
      colorlist[locationy+1][locationx-1]='#ffffff'
      sheet.getRange(locationy+1, locationx+1).setBackground('#ffffff')
      colorlist[locationy][locationx]='#ffffff'
      return(domove(sheet, turn, locationy+2, locationx-2, boxeslist, colorlist))
    }
  }
  //If none of the movement possibilities happened, return the position given
  if (flag==0) {return(sheet.getRange(locationy+1, locationx+1))}
}
