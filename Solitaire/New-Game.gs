function solitairenewgame() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Solitaire')

  var popup = Browser.msgBox('Are you sure you want to start a new game?', 'This will delete everything currently on the board', Browser.Buttons.YES_NO)

  if (popup=='yes') {

    //Get the current high score and stockpile draw style
    var high1 = sheet.getRange('D29').getValue()
    var high3 = sheet.getRange('D30').getValue()
    var drawstyle = sheet.getRange('V25').getValue()

    //Clear all of the current formatting and text
    sheet.getRange('A1:Z1000').breakApart().setBorder(false, false, false, false, false, false).clearContent().setHorizontalAlignment('left')

    //Set up stockpile card count & drawing style, move counter, and reset the high score display
    sheet.getRange('P1:Q1').merge().setValue('Cards left:')
    sheet.getRange('R1').setValue(24)
    sheet.getRange('V24:Z24').merge().setValue('Stockpile drawing style:').setHorizontalAlignment('center').setBorder(true, true, true, true, true, true)
    sheet.getRange('V25:Z25').merge().setHorizontalAlignment('center').setBorder(true, true, true, true, true, true, ).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['1 card at a time', '3 cards at a time'], true).build()).setValue(drawstyle)
    sheet.getRange('T1:V1').merge().setValue('Number of moves:')
    sheet.getRange('W1').setValue(0)
    sheet.getRange('A28:D28').merge().setValue('Fewest moves:').setHorizontalAlignment('center')
    sheet.getRange('A29:C29').merge().setValue('1 card at a time:')
    sheet.getRange('A30:C30').merge().setValue('3 cards at a time:')
    sheet.getRange('D29').setValue(high1)
    sheet.getRange('D30').setValue(high3)
    sheet.getRange('A28:D30').setBorder(true, true, true, true, true, true)

    //Color the board
    sheet.getRange('A1:O27').setBackground('#d9ead3').setFontSize(12)

    //Set up the foundations
    sheet.getRange('B1').setValue('♥').setFontColor('#ff0000').setHorizontalAlignment('center')
    sheet.getRange('D1').setValue('♦').setFontColor('#ff0000').setHorizontalAlignment('center')
    sheet.getRange('F1').setValue('♠').setFontColor('#000000').setHorizontalAlignment('center')
    sheet.getRange('H1').setValue('♣').setFontColor('#000000').setHorizontalAlignment('center')
    sheet.getRange('B2:B4').merge().setBorder(true, true, true, true, true, true)
    sheet.getRange('D2:D4').merge().setBorder(true, true, true, true, true, true)
    sheet.getRange('F2:F4').merge().setBorder(true, true, true, true, true, true)
    sheet.getRange('H2:H4').merge().setBorder(true, true, true, true, true, true)

    //Set up the stockpile
    if (drawstyle=='1 card at a time') {
      sheet.getRange('L2:L4').merge().setBorder(true, true, true, true, true, true)
    } else {
      sheet.getRange('L1:L3').merge().setBorder(true, true, true, true, true, true)
    }
    sheet.getRange('N2:N4').merge().setBorder(true, true, true, true, true, true).setBackground('#a4c2f4')

    //Set up the formatting of the lowest tableau cards
    sheet.getRange('B7:B9').merge().setBorder(true, true, true, true, true, true).setBackground('#ffffff')
    sheet.getRange('D8:D10').merge().setBorder(true, true, true, true, true, true).setBackground('#ffffff')
    sheet.getRange('F9:F11').merge().setBorder(true, true, true, true, true, true).setBackground('#ffffff')
    sheet.getRange('H10:H12').merge().setBorder(true, true, true, true, true, true).setBackground('#ffffff')
    sheet.getRange('J11:J13').merge().setBorder(true, true, true, true, true, true).setBackground('#ffffff')
    sheet.getRange('L12:L14').merge().setBorder(true, true, true, true, true, true).setBackground('#ffffff')
    sheet.getRange('N13:N15').merge().setBorder(true, true, true, true, true, true).setBackground('#ffffff')

    //Set up the formatting of the hidden tableau cards
    sheet.getRange('D7').setBorder(true, true, true, true, true, true).setBackground('#a4c2f4')
    sheet.getRange('F7:F8').setBorder(true, true, true, true, true, true).setBackground('#a4c2f4')
    sheet.getRange('H7:H9').setBorder(true, true, true, true, true, true).setBackground('#a4c2f4')
    sheet.getRange('J7:J10').setBorder(true, true, true, true, true, true).setBackground('#a4c2f4')
    sheet.getRange('L7:L11').setBorder(true, true, true, true, true, true).setBackground('#a4c2f4')
    sheet.getRange('N7:N12').setBorder(true, true, true, true, true, true).setBackground('#a4c2f4')

    //Now the randomness, generate an array of all the different cards and an array of matching suits
    var cards = [['A♥', '2♥', '3♥', '4♥', '5♥', '6♥', '7♥', '8♥', '9♥', '10♥', 'J♥', 'Q♥', 'K♥', 'A♦', '2♦', '3♦', '4♦', '5♦', '6♦', '7♦', '8♦', '9♦', '10♦', 'J♦', 'Q♦', 'K♦', 'A♠', '2♠', '3♠', '4♠', '5♠', '6♠', '7♠', '8♠', '9♠', '10♠', 'J♠', 'Q♠', 'K♠', 'A♣', '2♣', '3♣', '4♣', '5♣', '6♣', '7♣', '8♣', '9♣', '10♣', 'J♣', 'Q♣', 'K♣'], ['♥', '♥', '♥', '♥', '♥', '♥', '♥', '♥', '♥', '♥', '♥', '♥', '♥', '♦', '♦', '♦', '♦', '♦', '♦', '♦', '♦', '♦', '♦', '♦', '♦', '♦', '♠', '♠', '♠', '♠', '♠', '♠', '♠', '♠', '♠', '♠', '♠', '♠', '♠', '♣', '♣', '♣', '♣', '♣', '♣', '♣', '♣', '♣', '♣', '♣', '♣', '♣']]
  

    //Randomly shuffle the order of the cards
    cards = shuffle(cards)

    //Import this list into different places in the sheet, some are hidden so that they can be referenced later by other functions
    sheet.getRange('A800').setValue(drawstyle).setFontColor('#ffffff')
    var nextcolor = ''

    if(cards[1][0]=='♥' || cards[1][0]=='♦') {nextcolor='#ff0000'} else {nextcolor='#000000'}
    sheet.getRange('B7').setValue(cards[0][0]).setFontColor(nextcolor)

    if(cards[1][2]=='♥' || cards[1][2]=='♦') {nextcolor='#ff0000'} else {nextcolor='#000000'}
    sheet.getRange('D807').setValue(cards[0][1]).setFontColor('#ffffff')
    sheet.getRange('E807').setValue(cards[1][1]).setFontColor('#ffffff')
    sheet.getRange('D8').setValue(cards[0][2]).setFontColor(nextcolor)

    if(cards[1][5]=='♥' || cards[1][5]=='♦') {nextcolor='#ff0000'} else {nextcolor='#000000'}
    sheet.getRange('F807').setValue(cards[0][3]).setFontColor('#ffffff')
    sheet.getRange('G807').setValue(cards[1][3]).setFontColor('#ffffff')
    sheet.getRange('F808').setValue(cards[0][4]).setFontColor('#ffffff')
    sheet.getRange('G808').setValue(cards[1][4]).setFontColor('#ffffff')
    sheet.getRange('F9').setValue(cards[0][5]).setFontColor(nextcolor)

    if(cards[1][9]=='♥' || cards[1][9]=='♦') {nextcolor='#ff0000'} else {nextcolor='#000000'}
    sheet.getRange('H807').setValue(cards[0][6]).setFontColor('#ffffff')
    sheet.getRange('I807').setValue(cards[1][6]).setFontColor('#ffffff')
    sheet.getRange('H808').setValue(cards[0][7]).setFontColor('#ffffff')
    sheet.getRange('I808').setValue(cards[1][7]).setFontColor('#ffffff')
    sheet.getRange('H809').setValue(cards[0][8]).setFontColor('#ffffff')
    sheet.getRange('I809').setValue(cards[1][8]).setFontColor('#ffffff')
    sheet.getRange('H10').setValue(cards[0][9]).setFontColor(nextcolor)

    if(cards[1][14]=='♥' || cards[1][14]=='♦') {nextcolor='#ff0000'} else {nextcolor='#000000'}
    sheet.getRange('J807').setValue(cards[0][10]).setFontColor('#ffffff')
    sheet.getRange('K807').setValue(cards[1][10]).setFontColor('#ffffff')
    sheet.getRange('J808').setValue(cards[0][11]).setFontColor('#ffffff')
    sheet.getRange('K808').setValue(cards[1][11]).setFontColor('#ffffff')
    sheet.getRange('J809').setValue(cards[0][12]).setFontColor('#ffffff')
    sheet.getRange('K809').setValue(cards[1][12]).setFontColor('#ffffff')
    sheet.getRange('J810').setValue(cards[0][13]).setFontColor('#ffffff')
    sheet.getRange('K810').setValue(cards[1][13]).setFontColor('#ffffff')
    sheet.getRange('J11').setValue(cards[0][14]).setFontColor(nextcolor)

    if(cards[1][20]=='♥' || cards[1][20]=='♦') {nextcolor='#ff0000'} else {nextcolor='#000000'}
    sheet.getRange('L807').setValue(cards[0][15]).setFontColor('#ffffff')
    sheet.getRange('M807').setValue(cards[1][15]).setFontColor('#ffffff')
    sheet.getRange('L808').setValue(cards[0][16]).setFontColor('#ffffff')
    sheet.getRange('M808').setValue(cards[1][16]).setFontColor('#ffffff')
    sheet.getRange('L809').setValue(cards[0][17]).setFontColor('#ffffff')
    sheet.getRange('M809').setValue(cards[1][17]).setFontColor('#ffffff')
    sheet.getRange('L810').setValue(cards[0][18]).setFontColor('#ffffff')
    sheet.getRange('M810').setValue(cards[1][18]).setFontColor('#ffffff')
    sheet.getRange('L811').setValue(cards[0][19]).setFontColor('#ffffff')
    sheet.getRange('M811').setValue(cards[1][19]).setFontColor('#ffffff')
    sheet.getRange('L12').setValue(cards[0][20]).setFontColor(nextcolor)

    if(cards[1][27]=='♥' || cards[1][27]=='♦') {nextcolor='#ff0000'} else {nextcolor='#000000'}
    sheet.getRange('N807').setValue(cards[0][21]).setFontColor('#ffffff')
    sheet.getRange('O807').setValue(cards[1][21]).setFontColor('#ffffff')
    sheet.getRange('N808').setValue(cards[0][22]).setFontColor('#ffffff')
    sheet.getRange('O808').setValue(cards[1][22]).setFontColor('#ffffff')
    sheet.getRange('N809').setValue(cards[0][23]).setFontColor('#ffffff')
    sheet.getRange('O809').setValue(cards[1][23]).setFontColor('#ffffff')
    sheet.getRange('N810').setValue(cards[0][24]).setFontColor('#ffffff')
    sheet.getRange('O810').setValue(cards[1][24]).setFontColor('#ffffff')
    sheet.getRange('N811').setValue(cards[0][25]).setFontColor('#ffffff')
    sheet.getRange('O811').setValue(cards[1][25]).setFontColor('#ffffff')
    sheet.getRange('N812').setValue(cards[0][26]).setFontColor('#ffffff')
    sheet.getRange('O812').setValue(cards[1][26]).setFontColor('#ffffff')
    sheet.getRange('N13').setValue(cards[0][27]).setFontColor(nextcolor)

    //Finally put the list of stockpile cards in a hidden position
    for (i=28; i<52; i++) {
      sheet.getRange(800+i, 14).setValue(cards[0][i]).setFontColor('#ffffff')
      sheet.getRange(800+i, 15).setValue(cards[1][i]).setFontColor('#ffffff')
    }
  }
}

function shuffle(cards) {
  //Generate two empy arrays to be filled with cards and suits
  var newcards = [[], []]

  //Generate a random number that coresponds to some card in the initial array, append that card onto the empty array and remove that card from the initial array
  while (cards[0].length!=0) {
    var r = Math.floor(Math.random()*cards[0].length)
    newcards[0].push(cards[0][r])
    cards[0].splice(r, 1)
    newcards[1].push(cards[1][r])
    cards[1].splice(r, 1)
  }
  return(newcards)
}
