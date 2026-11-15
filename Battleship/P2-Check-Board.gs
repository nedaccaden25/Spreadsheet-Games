function battleshipP2checkboard() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Battleship Player 2')

  var boxes = sheet.getRange('O2:X11').getValues()

  var d = 0
  var s = 0
  var r = 0
  var b = 0
  var c = 0
  var up = 0
  var down = 0
  var left = 0
  var right = 0
  var counter = 0

  for (i=0; i<10; i++) {
    for (j=0; j<10; j++) {
      if (boxes[i][j]!='') {
        counter++
        up = checkup(boxes, i, j, boxes[i][j].toUpperCase())
        down = checkdown(boxes, i, j, boxes[i][j].toUpperCase())
        left = checkleft(boxes, i, j, boxes[i][j].toUpperCase())
        right = checkright(boxes, i, j, boxes[i][j].toUpperCase())
        if (boxes[i][j].toUpperCase()=='D' && (up+down==1 || left+right==1)) {d++}
        if (boxes[i][j].toUpperCase()=='S' && (up+down==2 || left+right==2)) {s++}
        if (boxes[i][j].toUpperCase()=='R' && (up+down==2 || left+right==2)) {r++}
        if (boxes[i][j].toUpperCase()=='B' && (up+down==3 || left+right==3)) {b++}
        if (boxes[i][j].toUpperCase()=='C' && (up+down==4 || left+right==4)) {c++}
      }
      up=0
      down=0
      left=0
      right=0
    }
  }

  if (c==5 && b==4 && r==3 && s==3 && d==2 && counter==17) {
    Browser.msgBox('Your board is valid', '', Browser.Buttons.OK)
    sheet.getRange('O2:X11').setBackground('#a4c2f4')
    for (i=0; i<10; i++) {
      for (j=0; j<10; j++) {
        if (boxes[i][j]!='') {sheet.getRange('O2').offset(i, j).setBackground('#cccccc')}
      }
    }
    sheet.getRange('O12').setBackground('#b6d7a8')
    SpreadsheetApp.getActive().getSheetByName('Battleship Player 1').getRange('B12').setBackground('#b6d7a8')
  } else {
    Browser.msgBox('Your ships are not correctly placed', 'Please make sure that you have set your board up correctly', Browser.Buttons.OK)
    return
  }
}

function checkup(boxes, i, j, checkagainst) {
  if (i==0) {return(0)}

  if (boxes[i-1][j].toUpperCase()==checkagainst) {return(1+checkup(boxes, i-1, j, checkagainst))}

  return(0)
}

function checkdown(boxes, i, j, checkagainst) {
  if (i==9) {return(0)}

  if (boxes[i+1][j].toUpperCase()==checkagainst) {return(1+checkdown(boxes, i+1, j, checkagainst))}

  return(0)
}

function checkleft(boxes, i, j, checkagainst) {
  if (j==0) {return(0)}

  if (boxes[i][j-1].toUpperCase()==checkagainst) {return(1+checkleft(boxes, i, j-1, checkagainst))}

  return(0)
}

function checkright(boxes, i, j, checkagainst) {
  if (j==9) {return(0)}

  if (boxes[i][j+1].toUpperCase()==checkagainst) {return(1+checkright(boxes, i, j+1, checkagainst))}

  return(0)
}
