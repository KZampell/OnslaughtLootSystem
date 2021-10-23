function updateSheet(sheetName){
  var mainSheet = BOOK.getSheetByName(sheetName)
  var curItem = mainSheet.getRange("A2").getCell(1,1)
  var outCell = curItem.offset(0, 2)
  while(curItem.getFontWeight() == 'bold'){
    curItem = curItem.offset(1, 0)
  }
  while (!curItem.getRichTextValue().getTextStyle().isItalic()){
    endTime = new Date().getTime()
    if (endTime - startTime >= 280000){
      Logger.log("Current Time: " + (endTime - startTime))
      break
    }
    if (curItem.getRichTextValue() == null){
      curItem = curItem.offset(1, 0)
    }
    else if (curItem.getRichTextValue().getTextStyle().isBold()){
      curItem = curItem.offset(1, 0)
    }
    else{
      var itemName = curItem.getValue().toString()
      var start = findMyIndex(itemName)
      var end = findFinalIndex(itemName, start)
      if (start >= 0){
        for (var i = start; i < end; i++){
          //Logger.log(TUPLES[i])
          var entry = TUPLES[i]
          if (itemName == entry.item){
            var outString = Utilities.formatString("%5.3f : %s", entry.score, entry.char)
            var teamLoc = mainSheet.getRange("A1:1").createTextFinder(entry.team).findNext()
            outCell = curItem.offset(0, teamLoc.getColumn()-1)
            if (outCell.getValue().toString().length == 0) {
              outCell.setBackground(entry.charClass)
              outCell.setValue(outString)
            }
            else {
              var tempS = outCell.getValue().toString()
              outCell.setValue(tempS + ", " + outString)
            }
          } 
        }
      }
      curItem.setFontWeight('bold')
      curItem = curItem.offset(1, 0)
    }
  }

  if (endTime - startTime < 280000){
    Logger.log("Current Time: " + (endTime - startTime) + ", Moving on to next phase")
    doneBool = true
  }
}