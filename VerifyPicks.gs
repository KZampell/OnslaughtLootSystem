function parsePicks(){
  var charSheet = BOOK.getSheetByName(CHAR_POINT)
  var ptr = 50
  var bracketPt = 0
  var resets = [50, 47, 44, 41]
  var curChar = charSheet.getRange("A3")
  while (curChar.getValue().toString().length > 1){
    var curCell = curChar.offset(0, 1)
    ptr = 50
    while (ptr > 38){
      if (resets.includes(ptr)){
        bracketPt = 0
      }
      var checkTxt = curCell.getFontColor()
      if (checkTxt == "#ff0000"){
        bracketPt++
        var followUp = curCell.offset(1, 0).getValue().toString()
        if (followUp.length > 1){
          curChar.setFontColor("#00ff00")
          break
        }
      }
      else if (checkTxt == "#0000ff"){
        bracketPt++
      }
      if (bracketPt > 3){
        curChar.setFontColor("#00ff00")
        break
      }
      curCell = curCell.offset(0, 1)
      ptr--
    }
    curChar = curChar.offset(2, 0)
  }
}