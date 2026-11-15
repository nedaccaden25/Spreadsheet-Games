function snakeGame() {
  
  const sheet = SpreadsheetApp.getActive().getSheetByName('Snake')
  //Gets which page in the spreadsheet that these changes will be done on

  var score = 0
  //Variable used when displaying the score

  sheet.getRange('A19').setValue(0)
  //Reset the score on the sheet

  var snakelength = 3
  //Sets the initial length of the snake

  var snakedirection = 'Right'
  //Sets the initial direction the snake is facing

  var applepos = sheet.getRange('M8')
  //Sets the initial position of the apple

  var togrow = 0
  //Sets up a flag for later use

  var dead = 0
  //Sets up a flag for later use

  let snake = new Array(255);
  //Create an array to store the individual locations of each segment

  snake[0] = sheet.getRange('D8')
  snake[1] = sheet.getRange('C8')
  snake[2] = sheet.getRange('B8')

  sheet.getRange('A1:Q15').setBackground('#ffffff').setValue('')
  //Clears the board's colors and text

  applepos.setBackground('#e06666')
  //Colors the apple's position

  snake[0].setBackground('#93c47d')
  snake[1].setBackground('#93c47d')
  snake[2].setBackground('#93c47d')
  //Colors the Snake's cells

  SpreadsheetApp.flush()
  //Applies the changes

  while (dead == 0) {

    snakedirection = setSnakeDirection(snakedirection, sheet)
    //Set the snake's direction to the selected cell

    if (togrow == 1) {
      snakelength = snakelength+1
    }
    //If the snake needs to grow this movement, update the length and scores

    snake = moveSnake(snake, snakedirection, snakelength, togrow)
    //Move the snake and check for wall collisions

    togrow = 0

    if (snake[0]!=0){
      dead = selfCollisionCheck(snake, snakelength)
      //Check for collisions with the body
    }

    if (snake[0]==0) {
      dead = 1
      continue
      //If a wall was hit, end the game with a death screen
    }

    if (snake[0].getColumn()==applepos.getColumn() && snake[0].getRow()==applepos.getRow()) {
      applepos = randomizeApple(snake, snakelength, sheet)
      score = updateScores(score, sheet)
      togrow=1
      //If the snake head is on the apple, place a new apple, update the scores, and set a flag that the snake needs to grow the next time it moves
    }

    SpreadsheetApp.flush()
    //Applies the changes

    Utilities.sleep(250)
    //Wait for a quarter of a second
  }

  sheet.getRange('H7').setValue('Y')
  sheet.getRange('I7').setValue('O')
  sheet.getRange('J7').setValue('U')
  sheet.getRange('G9').setValue('D')
  sheet.getRange('H9').setValue('I')
  sheet.getRange('I9').setValue('E')
  sheet.getRange('J9').setValue('D')
  sheet.getRange('K9').setValue('!')
  sheet.getRange('G6:K10').setBackground('#cccccc')
  //Once the main loop is escaped, that means you died, so display the death screen
}

function moveSnake(snake, snakedirection, snakelength, togrow) {

  if (togrow == 0) {
    snake[snakelength-1].setBackground('#ffffff')
    //Set the end of the snake to white only if the snake doesn't need to grow
  }

  for (i=snakelength-1; i!=0; i--) {
    snake[i]=snake[i-1]
    //Move each segment of the snake forward except the head
  }

  if (snakedirection == 'Right') {

    if (snake[0].getColumn()==17) {
      snake[0]=0
      //Check for wall collision

    } else {
      snake[0]=snake[0].offset(0,1)
      //Move the head right one

    }
  } else if (snakedirection == 'Up') {

    if (snake[0].getRow()==1) {
      snake[0]=0
      //Check for wall collision

    } else {
      snake[0]=snake[0].offset(-1,0)
      //Move the head up one

    }
  } else if (snakedirection == 'Left') {

    if (snake[0].getColumn()==1) {
      snake[0]=0
      //Check for wall collision

    } else {
      snake[0]=snake[0].offset(0,-1)
      //Move the head left one

    }
  } else if (snakedirection == 'Down') {

    if (snake[0].getRow()==15) {
      snake[0]=0
      //Check for wall collision

    } else {
      snake[0]=snake[0].offset(1,0)
      //Move the head down one

    }
  }
  if (snake[0]!=0) {
    snake[0].setBackground('#93c47d')
    //Set the new head space green
  }

  return(snake)
  //Update the array
}

function setSnakeDirection(snakedirection, sheet) {

  var rightbox = sheet.getRange('O20').isChecked()
  if (rightbox) {sheet.getRange('O20').uncheck()}
  var upbox = sheet.getRange('M18').isChecked()
  if (upbox) {sheet.getRange('M18').uncheck()}
  var leftbox = sheet.getRange('K20').isChecked()
  if (leftbox) {sheet.getRange('K20').uncheck()}
  var downbox = sheet.getRange('M20').isChecked()
  if (downbox) {sheet.getRange('M20').uncheck()}
  //Gets the status of the check boxes while unchecking them

  if (rightbox && snakedirection!='Left') {
    return('Right')
    //Set the direction to right

  } else if (upbox && snakedirection!='Down') {
    return('Up')
    //Set the direction to up
      
  } else if (leftbox && snakedirection!='Right') {
    return('Left')
    //Set the direction to left
      
  } else if (downbox && snakedirection!='Up') {
    return('Down')
    //Set the direction to down

  } else {
    return(snakedirection)
    //If no boxes are set, or you tried to turn around, do nothing
  }
}

function selfCollisionCheck(snake, snakelength){

  for (i=snakelength-1; i!=0; i--){
    if (snake[0].getColumn()==snake[i].getColumn() && snake[0].getRow()==snake[i].getRow()) {
      return(1)
      //If the head overlaps with the body somewhere, mark the collision
    }
  }
  return(0)
  //If nothing overlapped, return without changing anything
}

function randomizeApple (snake, snakelength, sheet) {

  var posfound = 0
  //Sets up a flag for later use

  while (posfound == 0) {

    var newpos = sheet.getRange(Math.floor(Math.random() * 14)+1, Math.floor(Math.random() * 16)+1)
    //Generates a random cell on the board

    posfound = 1
    //If the randomly generated value is not on the snake, this value will remain, otherwise it will get overwritten

    for (i=snakelength-1; i!=-1; i--) {
      posfound = 1
      if (newpos.getColumn()==snake[i].getColumn() && newpos.getRow()==snake[i].getRow()) {
        posfound = 0
        break
        //If the randomly generated value was on the snake, reset the flag and get out of the loop
      }
    }
  }

  newpos.setBackground('#e06666')
  return(newpos)
  //Turn the new apple cell red and output the new location
}

function updateScores(score, sheet) {
  score = score+1
  sheet.getRange('A19').setValue(score)
  //Update the score

  if (score>sheet.getRange('A21').getValue()) {
    sheet.getRange('A21').setValue(score)
  }
  //Update the high score if applicable

  return(score)
}
