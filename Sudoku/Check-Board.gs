function sudokucheckboard() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Sudoku')

  //Initialize a 3D array to store the locations of each cell to reference more easily later
  var boardranges = [
    [
      [sheet.getRange('A1'), sheet.getRange('B1'), sheet.getRange('C1')],
      [sheet.getRange('A2'), sheet.getRange('B2'), sheet.getRange('C2')],
      [sheet.getRange('A3'), sheet.getRange('B3'), sheet.getRange('C3')]
    ],
    [
      [sheet.getRange('D1'), sheet.getRange('E1'), sheet.getRange('F1')],
      [sheet.getRange('D2'), sheet.getRange('E2'), sheet.getRange('F2')],
      [sheet.getRange('D3'), sheet.getRange('E3'), sheet.getRange('F3')]
    ],
    [
      [sheet.getRange('G1'), sheet.getRange('H1'), sheet.getRange('I1')],
      [sheet.getRange('G2'), sheet.getRange('H2'), sheet.getRange('I2')],
      [sheet.getRange('G3'), sheet.getRange('H3'), sheet.getRange('I3')]
    ],
    [
      [sheet.getRange('A4'), sheet.getRange('B4'), sheet.getRange('C4')],
      [sheet.getRange('A5'), sheet.getRange('B5'), sheet.getRange('C5')],
      [sheet.getRange('A6'), sheet.getRange('B6'), sheet.getRange('C6')]
    ],
    [
      [sheet.getRange('D4'), sheet.getRange('E4'), sheet.getRange('F4')],
      [sheet.getRange('D5'), sheet.getRange('E5'), sheet.getRange('F5')],
      [sheet.getRange('D6'), sheet.getRange('E6'), sheet.getRange('F6')]
    ],
    [
      [sheet.getRange('G4'), sheet.getRange('H4'), sheet.getRange('I4')],
      [sheet.getRange('G5'), sheet.getRange('H5'), sheet.getRange('I5')],
      [sheet.getRange('G6'), sheet.getRange('H6'), sheet.getRange('I6')]
    ],
    [
      [sheet.getRange('A7'), sheet.getRange('B7'), sheet.getRange('C7')],
      [sheet.getRange('A8'), sheet.getRange('B8'), sheet.getRange('C8')],
      [sheet.getRange('A9'), sheet.getRange('B9'), sheet.getRange('C9')]
    ],
    [
      [sheet.getRange('D7'), sheet.getRange('E7'), sheet.getRange('F7')],
      [sheet.getRange('D8'), sheet.getRange('E8'), sheet.getRange('F8')],
      [sheet.getRange('D9'), sheet.getRange('E9'), sheet.getRange('F9')]
    ],
    [
      [sheet.getRange('G7'), sheet.getRange('H7'), sheet.getRange('I7')],
      [sheet.getRange('G8'), sheet.getRange('H8'), sheet.getRange('I8')],
      [sheet.getRange('G9'), sheet.getRange('H9'), sheet.getRange('I9')]
    ]
  ]

  //Get the current values on the board to reference more easily later
  var board = [
    [
      [sheet.getRange('A1').getValue(), sheet.getRange('B1').getValue(), sheet.getRange('C1').getValue()],
      [sheet.getRange('A2').getValue(), sheet.getRange('B2').getValue(), sheet.getRange('C2').getValue()],
      [sheet.getRange('A3').getValue(), sheet.getRange('B3').getValue(), sheet.getRange('C3').getValue()]
    ],
    [
      [sheet.getRange('D1').getValue(), sheet.getRange('E1').getValue(), sheet.getRange('F1').getValue()],
      [sheet.getRange('D2').getValue(), sheet.getRange('E2').getValue(), sheet.getRange('F2').getValue()],
      [sheet.getRange('D3').getValue(), sheet.getRange('E3').getValue(), sheet.getRange('F3').getValue()]
    ],
    [
      [sheet.getRange('G1').getValue(), sheet.getRange('H1').getValue(), sheet.getRange('I1').getValue()],
      [sheet.getRange('G2').getValue(), sheet.getRange('H2').getValue(), sheet.getRange('I2').getValue()],
      [sheet.getRange('G3').getValue(), sheet.getRange('H3').getValue(), sheet.getRange('I3').getValue()]
    ],
    [
      [sheet.getRange('A4').getValue(), sheet.getRange('B4').getValue(), sheet.getRange('C4').getValue()],
      [sheet.getRange('A5').getValue(), sheet.getRange('B5').getValue(), sheet.getRange('C5').getValue()],
      [sheet.getRange('A6').getValue(), sheet.getRange('B6').getValue(), sheet.getRange('C6').getValue()]
    ],
    [
      [sheet.getRange('D4').getValue(), sheet.getRange('E4').getValue(), sheet.getRange('F4').getValue()],
      [sheet.getRange('D5').getValue(), sheet.getRange('E5').getValue(), sheet.getRange('F5').getValue()],
      [sheet.getRange('D6').getValue(), sheet.getRange('E6').getValue(), sheet.getRange('F6').getValue()]
    ],
    [
      [sheet.getRange('G4').getValue(), sheet.getRange('H4').getValue(), sheet.getRange('I4').getValue()],
      [sheet.getRange('G5').getValue(), sheet.getRange('H5').getValue(), sheet.getRange('I5').getValue()],
      [sheet.getRange('G6').getValue(), sheet.getRange('H6').getValue(), sheet.getRange('I6').getValue()]
    ],
    [
      [sheet.getRange('A7').getValue(), sheet.getRange('B7').getValue(), sheet.getRange('C7').getValue()],
      [sheet.getRange('A8').getValue(), sheet.getRange('B8').getValue(), sheet.getRange('C8').getValue()],
      [sheet.getRange('A9').getValue(), sheet.getRange('B9').getValue(), sheet.getRange('C9').getValue()]
    ],
    [
      [sheet.getRange('D7').getValue(), sheet.getRange('E7').getValue(), sheet.getRange('F7').getValue()],
      [sheet.getRange('D8').getValue(), sheet.getRange('E8').getValue(), sheet.getRange('F8').getValue()],
      [sheet.getRange('D9').getValue(), sheet.getRange('E9').getValue(), sheet.getRange('F9').getValue()]
    ],
    [
      [sheet.getRange('G7').getValue(), sheet.getRange('H7').getValue(), sheet.getRange('I7').getValue()],
      [sheet.getRange('G8').getValue(), sheet.getRange('H8').getValue(), sheet.getRange('I8').getValue()],
      [sheet.getRange('G9').getValue(), sheet.getRange('H9').getValue(), sheet.getRange('I9').getValue()]
    ]
  ]

  var flag=0
  var failed=0

  for (i=0; i<9; i++) {
    for (j=0; j<3; j++) {
      for (k=0; k<3; k++) {
        if (board[i][j][k]!='' && board[i][j][k]!=0) {

          //Check the box 3x3 box for the same number
          for (a=0; a<3; a++) {
            for (b=0; b<3; b++) {
              if (board[i][a][b]==board[i][j][k] && !(a==j && b==k)) {
                failed=1
                boardranges[i][j][k].setBackground('#ea9999')
                break
              }
            }
          }

          //Check the row for the same number
          for (a=0; a<3; a++) {
            for (b=0; b<3; b++) {
              if (board[Math.floor(i/3)*3+a][j][b]==board[i][j][k] && !(Math.floor(i/3)*3+a==i && b==k)) {
                failed=1
                boardranges[i][j][k].setBackground('#ea9999')
                break
              }
            }
          }

          //Check the column for the same number
          for (a=0; a<3; a++) {
            for (b=0; b<3; b++) {
              if (board[i%3+a*3][b][k]==board[i][j][k] && !(i%3+a*3==i && b==j)) {
                failed=1
                boardranges[i][j][k].setBackground('#ea9999')
                break
              }
            }
          }
        } else {flag=1}
      }
    }
  }

  //If no spaces were blank, and every space was valid, display that the board is complete
  if (failed==0 && flag==0) {
    Browser.msgBox('Congratulations!', 'Your board is complete!', Browser.Buttons.OK)
  
  //If there was at least one blank space, but every filled space was valid, display that there are no errors
  } else if (failed==0 && flag==1) {
    Browser.msgBox('The board is correct so far', 'There are no errors on the board, but not every space is filled yet', Browser.Buttons.OK)

  //If there was at least one invalid space, diasplay that there was an error
  } else if (failed==1) {
    Browser.msgBox('Uh Oh!', 'At least one of your spaces is invalid, the conflicting spaces have been colored red', Browser.Buttons.OK)
  }
}
