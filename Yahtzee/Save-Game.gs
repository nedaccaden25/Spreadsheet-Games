function yahtzeesave() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Yahtzee')
  const savesheet = SpreadsheetApp.getActive().getSheetByName('Saved Yahtzee Boards')

  var offset = sheet.getRange('V22').getValue()-1

  var resetyn = Browser.msgBox('Are you sure you want to save the current game to save slot '+(offset+1), 'whatever is currently in that save slot will be lost', Browser.Buttons.YES_NO)
  if (resetyn == 'yes') {

    var boardvals = sheet.getRange('X1:AC22').getValues()
    var boardcolors = sheet.getRange('X1:AC22').getBackgrounds()
    var boardtextcolors = sheet.getRange('X1:AC22').getFontColors()
    var boardtextsizes = sheet.getRange('X1:AC22').getFontSizes()

    savesheet.getRange('F5:F11').offset(0, offset).removeCheckboxes()
    savesheet.getRange('A3:F24').offset(0, offset*7).setValues(boardvals)
    savesheet.getRange('A3:F24').offset(0, offset*7).setBackgrounds(boardcolors)
    savesheet.getRange('A3:F24').offset(0, offset*7).setFontColors(boardtextcolors)
    savesheet.getRange('A3:F24').offset(0, offset*7).setFontSizes(boardtextsizes)

    SpreadsheetApp.flush()

    if (!savesheet.getRange('F5').offset(0, offset*7).isBlank() && (savesheet.getRange('F5').offset(0, offset*7).getValue()===false || savesheet.getRange('F5').offset(0, offset*7).getValue()===true)) {
      savesheet.getRange('F5').offset(0, offset*7).insertCheckboxes()
    }
    if (!savesheet.getRange('F8').offset(0, offset*7).isBlank() && (savesheet.getRange('F8').offset(0, offset*7).getValue()===false || savesheet.getRange('F8').offset(0, offset*7).getValue()===true)) {
      savesheet.getRange('F8').offset(0, offset*7).insertCheckboxes()
    }
    if (!savesheet.getRange('F11').offset(0, offset*7).isBlank() && (savesheet.getRange('F11').offset(0, offset*7).getValue()===false || savesheet.getRange('F11').offset(0, offset*7).getValue()===true)) {
      savesheet.getRange('F11').offset(0, offset*7).insertCheckboxes()
    }
  }
}
