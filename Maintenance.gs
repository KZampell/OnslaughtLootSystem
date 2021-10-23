function wipeSlate(sheetName){
  var shet = BOOK.getSheetByName(sheetName)
  var activeRange = shet.getRange("C3:3")
  var pointText = activeRange.getCell(1, 1).offset(0, -2)
  while (pointText.getRichTextValue().getTextStyle().isItalic() != true){
    if (pointText.getRichTextValue().getText().toString().length > 1){
      if (activeRange.getCell(1,1).getBackground() != "white"){
        activeRange.setBackground("white")
        activeRange.clearContent()
      }
    }
    activeRange = activeRange.offset(1, 0)
    pointText = activeRange.getCell(1, 1).offset(0, -2)
  }
}

function colorLists(){
  var shet = BOOK.getSheetByName(CHAR_POINT)
  var resSheet = BOOK.getSheetByName(RES_IT)
  var limSheet = BOOK.getSheetByName(LIM_IT)
  var unSheet = BOOK.getSheetByName(UN_IT)
  var activeRange = shet.getRange("A3:Z4")
  var nameCell = activeRange.getCell(1,1)
  var curCell = activeRange.getCell(1,2)
  var ptsCount = 50
  while (nameCell.getRichTextValue().getText().toString().length > 0){
    if (activeRange.getCell(1, 1).getFontWeight() != "bold"){
      while (ptsCount > 25){
        var itemName = curCell.getValue()
        if(itemName.toString().length > 1){
          var resFind = resSheet.createTextFinder(itemName.toString())
          resFind.matchEntireCell(true)
          var loc = resFind.findNext()
          if (loc != null){
            curCell.setFontColor("#ff0000")
          }
          else {
            var limFind = limSheet.createTextFinder(itemName.toString())
            limFind.matchEntireCell(true)
            loc = limFind.findNext()
            if (loc != null){
              curCell.setFontColor("#0000ff")
            }
            else {
              var unFind = unSheet.createTextFinder(itemName.toString())
              unFind.matchEntireCell(true)
              loc = unFind.findNext()
              if (loc == null){
                curCell.setFontColor("#00ff00")
              }
              else {
                curCell.setFontColor("#000000")
              }
            }
          }
        }
        var topColor = curCell.getFontColor()
        curCell = curCell.offset(1, 0)
        itemName = curCell.getValue()
        itemColor = curCell.getFontColor()
        if (topColor == "#ff0000"){
          curCell.setFontColor("#00ff00")
        }
        else if(itemName.toString().length > 1){
          var resFind = resSheet.createTextFinder(itemName.toString())
          resFind.matchEntireCell(true)
          var loc = resFind.findNext()
          if (loc != null){
            curCell.setFontColor("#ff0000")
          }
          else {
            var limFind = limSheet.createTextFinder(itemName.toString())
            limFind.matchEntireCell(true)
            loc = limFind.findNext()
            if (loc != null){
              curCell.setFontColor("#0000ff")
            }
            else {
              var unFind = unSheet.createTextFinder(itemName.toString())
              unFind.matchEntireCell(true)
              loc = unFind.findNext()
              if (loc == null){
                curCell.setFontColor("#00ff00")
              }
              else {
                curCell.setFontColor("#000000")
              }
            }
          }
        }
        
        curCell = curCell.offset(-1, 1)
        ptsCount--
      }
    }
    nameCell.setFontWeight("bold")
    activeRange = activeRange.offset(2, 0)
    nameCell = activeRange.getCell(1,1)
    curCell = activeRange.getCell(1,2)
    ptsCount = 50
      
  }
}

function firstAttendance(){
  var srcSheet = BOOK.getSheetByName(CHAR_POINT)
  var attSheet = BOOK.getSheetByName(ATT)
  var range = srcSheet.getRange("A3")
  var cName = range.getValue().toString()
  var splat = attSheet.getRange("A2")
  while (cName.length > 1){
    splat.setValue(cName)
    range = range.offset(2, 0)
    splat = splat.offset(1, 0)
    cName = range.getValue().toString()
  }
}