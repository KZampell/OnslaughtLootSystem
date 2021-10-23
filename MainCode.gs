var CHAR_POINT = "CharPointAlloc"
var RES_IT = "ReservedItems"
var LIM_IT = "LimitedItems"
var UN_IT = "Unlimited"
var ATT = "Attendance"
var PASS = "Passes/Ties"
var DAT = "DataStore"
var TEAMS = "Teams"
var LOOT = "LootLog"
var P1 = "P1"
var P2 = "P2"
var P3 = "P3"
var BOOK = SpreadsheetApp.openById("Sheet ID Goes Here")
var TUPLES = []
var startTime = 0
var endTime = 0

var adminList = ["user1@email.com", "user2@email.net"]


function updateItemValues(){
  var charSheet = BOOK.getSheetByName(CHAR_POINT)
  var storage = BOOK.getSheetByName(DAT)

  var activeRange = charSheet.getRange("A3:Z4")
  var storageRange = storage.getRange("A1:E1")
  while(storageRange.getCell(1,1).getValue().toString().length > 0 & storageRange.getCell(1,1).getFontStyle() != 'italic'){
    storageRange = storageRange.offset(1, 0)
  }

  while (activeRange.getCell(1,1).getRichTextValue().getText().toString().length > 0){
    endTime = new Date().getTime()
    if (endTime - startTime >= 280000){
      Logger.log("Current Time: " + (endTime - startTime))
      break
    }
    var ptsVal = 50
    var charName = activeRange.getCell(1,1).getValue()
    var charClass = activeRange.getCell(1,1).getBackground()
    var specAndAlt = activeRange.getCell(2,1).getRichTextValue().getText().toLowerCase().toString()
    var teamName = findTeams(charName)
    if (teamName == ""){
      activeRange.getCell(1,1).setFontWeight('bold')
      activeRange = activeRange.offset(2, 0)
      continue
    }
    var readCell = activeRange.getCell(1,2)
    if(activeRange.getCell(1,1).getFontWeight() != 'bold'){
      while (ptsVal > 25){
        var activeText = readCell.getRichTextValue().getText().toString()
        if (activeText.length > 1){
          //updateTuple(readCell, ptsVal, charName, specAndAlt, charClass, teamName)
          var skip = pasteData(storageRange, readCell, ptsVal, specAndAlt, charName, charClass, teamName)
          if (skip){
            storageRange = storageRange.offset(1, 0)
          }
        }
        readCell = readCell.offset(1, 0)
        activeText = readCell.getRichTextValue().getText().toString()
        if(activeText.length > 1){
          //updateTuple(readCell, ptsVal, charName, specAndAlt, charClass, teamName)
          var skip = pasteData(storageRange, readCell, ptsVal, specAndAlt, charName, charClass, teamName)
          if (skip) {
            storageRange = storageRange.offset(1, 0)
          }
        }
        readCell = readCell.offset(-1, 0)
        readCell = readCell.offset(0, 1)
        ptsVal -= 1
      }
      activeRange.getCell(1,1).setFontWeight('bold')
      Logger.log("Completed " + charName)
    }
    
    activeRange = activeRange.offset(2, 0)
  }
  Logger.log("Completed parse")

  if (endTime - startTime < 280000){
    Logger.log("Current Time: " + (endTime - startTime) + ", Moving on to next phase")
    doneBool = true
  }

  if(storageRange.getCell(1,1).getFontStyle() != 'italic'){
    storageRange.getCell(1,1).setValue("End of Line")
    storageRange.getCell(1,1).setFontStyle("italic")
  }
}

function pasteData(storageRange, tarCell, pts, specAndAlt, char, charClass, team){
  var storage = BOOK.getSheetByName(DAT)
  var color = tarCell.getFontColor().toString()
  if (color == "#ff0000"){
    var charSheet = BOOK.getSheetByName(RES_IT)
  }
  else if (color == "#0000ff"){
    var charSheet = BOOK.getSheetByName(LIM_IT)
  }
  else{
    var charSheet = BOOK.getSheetByName(UN_IT)
  }
  var sName = tarCell.getValue().toString()
  //var alt = (specAndAlt.indexOf("alt") > -1)
  var spec = specAndAlt.substring(0, 4)
  var [atten, playerRaids] = fetchAttendance(char)
  var tFind = charSheet.createTextFinder(sName)
  tFind.matchEntireCell(true)
  var loc = tFind.findNext()
  if (loc != null){
    if (loc.offset(0, 3).getRichTextValue().getText().toString().length > 1){
      var tarSpec = loc.offset(0, 3).getRichTextValue().getText().toLowerCase().toString()
      var specBool = (spec == tarSpec) || (tarSpec == "all")
      if (spec == "tank" && tarSpec == "cast" && charClass == "#ea9999"){
        specBool = true
      } else if (spec == "tank" && tarSpec == "phys" && (charClass == "#ff9900" || charClass == "#7f6000")){
        specBool = true
      }
    }
    else {
      //Default to assume mainspec
      specBool = true
    }
  }
  else{
    //Full match did not work, try for a partial
    var xFind = charSheet.createTextFinder(sName)
    loc = xFind.findNext()
    if (loc != null){

      if (loc.offset(0, 3).getRichTextValue().getText().toString().length > 1){
        var tarSpec = loc.offset(0, 3).getRichTextValue().getText().toLowerCase().toString()
        var specBool = (spec == tarSpec) || (tarSpec == "all")
      }
      else {
        //Default to assume mainspec
        specBool = true
      }
    } else {
      var bonus = 0
      var specBool = false
      Logger.log("I could not find " + sName)
    }
  }
  if (atten > 0 && loc != null){
    var drops = fetchDrops(sName, loc, playerRaids, charSheet.getRange(3, 7))
    var bonus = calcBonus(atten, drops)
  } else {
    var drops = 0
    var bonus = 0
  }
  var iName = tarCell.getRichTextValue().getText().toString()
  var specBonus = bonus + countTies(iName, char)
  if (specBool == false){
    var scoreVal = (0.9*pts) + specBonus
  }
  else{
    var scoreVal = pts + specBonus
  }
  if(loc != null){
    var raid = loc.offset(0, 4).getValue().toString()
  } else {
    raid = "NA"
  }
  var raidArray = [{}];
  raidArray = team.split(", ");
  for (str in raidArray){
    var loca = raidArray[str].indexOf(raid)
    if (loca > -1){
      if(storageRange.getCell(1,1).getFontStyle() == 'italic'){
        storage.insertRowBefore(storageRange.getRow())
      }
      var itemCell = storageRange.getCell(1,1)
      itemCell.setFontStyle('normal')
      var scoreCell = storageRange.getCell(1, 2)
      var nameCell = storageRange.getCell(1,3)
      var classCell = storageRange.getCell(1,4)
      var teamCell = storageRange.getCell(1,5)
      itemCell.setValue(iName)
      scoreCell.setValue(scoreVal.toString())
      nameCell.setValue(char)
      classCell.setValue(charClass)
      teamCell.setValue(raidArray[str].charAt(0))
      return true
    }
  }
}


function updateTuple(tarCell, pts, cName, specAndAlt, charClass, teamN){
  var color = tarCell.getFontColor().toString()
  if (color == "#ff0000"){
    var charSheet = BOOK.getSheetByName(RES_IT)
  }
  else if (color == "#0000ff"){
    var charSheet = BOOK.getSheetByName(LIM_IT)
  }
  else{
    var charSheet = BOOK.getSheetByName(UN_IT)
  }
  var sName = tarCell.getValue().toString()
  var alt = (specAndAlt.indexOf("alt") > -1)
  var spec = specAndAlt.substring(0, 4)
  var tFind = charSheet.createTextFinder(sName)
  tFind.matchEntireCell(true)
  var loc = tFind.findNext()
  if (loc != null){
    var [atten, playerRaids] = fetchAttendance(cName)
    var drops = fetchDrops(sName, loc, playerRaids, charSheet.getRange(3, 7))
    var bonus = calcBonus(atten, drops)
    if (loc.offset(0, 3).getRichTextValue().getText().toString().length > 1){
      var tarSpec = loc.offset(0, 3).getRichTextValue().getText().toLowerCase().toString()
      var specBool = (spec == tarSpec) || (tarSpec == "all")
    }
    else {
      //Default to assume mainspec
      specBool = true
    }
  }
  else{
    var xFind = charSheet.createTextFinder(sName)
    loc = xFind.findNext()
    if (loc != null){
      var [atten, playerRaids] = fetchAttendance(cName)
      var drops = fetchDrops(sName, loc, playerRaids, charSheet.getRange(3, 7))
      var bonus = calcBonus(atten, drops)
      if (loc.offset(0, 3).getRichTextValue().getText().toString().length > 1){
        var tarSpec = loc.offset(0, 3).getRichTextValue().getText().toLowerCase().toString()
        var specBool = (spec == tarSpec) || (tarSpec == "all")
      }
      else {
        //Default to assume mainspec
        specBool = true
      }
    } else {
      var bonus = 0
      var specBool = false
      Logger.log("I could not find " + sName)
    }
  }
  var iName = tarCell.getRichTextValue().getText().toString()
  var specBonus = bonus + countTies(iName, cName)
  if (specBool == false){
    var scoreVal = (0.9*pts) + specBonus
  }
  else{
    var scoreVal = pts + specBonus
  }
  if(loc != null){
    var raid = loc.offset(0, 4).getValue().toString()
  } else {
    raid = "NA"
  }
  var raidArray = [{}];
  raidArray = teamN.split(", ");
  for (str in raidArray){
    var loca = raidArray[str].indexOf(raid)
    if (loca > -1){
      TUPLES.push(new tEntry(iName, scoreVal, cName, charClass, raidArray[str].charAt(0)))
    }
  }
}


function findMyIndex(n){
  var idx = 0
  var ret = -1
  while(ret < 0 && idx < TUPLES.length){
    var obj = TUPLES[idx]
    if (obj.item == n){
      ret = idx
      return ret
    }
    else{
      idx++
    }
  }
}

function findFinalIndex(n, i){
  var ret = -1
  var idx = i
  while (ret < 0 && idx < TUPLES.length){
    var obj = TUPLES[idx]
    if (obj.item != n){
      ret = idx
      return ret
    }
    else{
      idx++
    }
  }
  return TUPLES.length
}
