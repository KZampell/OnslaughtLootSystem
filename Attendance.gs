function parseAttenders() {
  var teamSheet = BOOK.getSheetByName(TEAMS)
  var attSheet = BOOK.getSheetByName(ATT)
  var resSheet = BOOK.getSheetByName(RES_IT)
  var limSheet = BOOK.getSheetByName(LIM_IT)
  var unSheet = BOOK.getSheetByName(UN_IT)

  var raidDate1 = teamSheet.getRange("A1")
  var raidDate2 = teamSheet.getRange("B1")
  var exists = false
    
  dateCell = resSheet.getRange("G3")
  while (dateCell.getValue().toString().length > 0){
    if (dateCell.getValue().toString() == raidDate1.getValue().toString()){
      exists = true
      break
    }
    dateCell = dateCell.offset(0, 1)
  }
  if (exists == false){
    dateCell.setValue(raidDate1.getValue())
  }

  exists = false
  dateCell = limSheet.getRange("G3")
  while (dateCell.getValue().toString().length > 0){
    if (dateCell.getValue().toString() == raidDate.getValue().toString()){
      exists = true
      break
    }
    dateCell = dateCell.offset(0, 1)
  }
  if (exists == false){
    dateCell.setValue(raidDate.getValue())
  }

  exists = false
  dateCell = unSheet.getRange("G3")
  while (dateCell.getValue().toString().length > 0){
    if (dateCell.getValue().toString() == raidDate.getValue().toString()){
      exists = true
      break
    }
    dateCell = dateCell.offset(0, 1)
  }
  if (exists == false){
    dateCell.setValue(raidDate.getValue())
  }

  exists = false
  var dateCell = attSheet.getRange("F1").getCell(1,1)
  if (dateCell.getValue().toString().length > 0){
    attSheet.insertColumnBefore(6)
    //dateCell = dateCell.offset(0, -1)
  }
  dateCell.setValue(raidDate.getValue())

  var dateCol = dateCell.getColumn()
  var teamRange = teamSheet.getRange("A2:A")
  var curTeam = teamRange.getCell(1,1)
  
  while (curTeam.getRichTextValue().getText().length > 0){
    var personTracker = curTeam.offset(1, 0)
    var teamN = curTeam.getValue().toString().charAt(0)
    while (personTracker.getRichTextValue().getText().length > 0){
      var charN = personTracker.getValue().toString()
      var attBool = (personTracker.getBackground() == "#00ff00")
      if (attBool){
        var tFind = attSheet.createTextFinder(charN)
        tFind.matchEntireCell(true)
        var loc = tFind.findNext()
        if (loc != null){
          var tarCell = loc.offset(0, dateCol-1)
          if (tarCell.getRichTextValue().getText().toString().length == 0){
            tarCell.setValue(teamN)
          }
          else {
            var tmp = tarCell.getValue()
            tmp = tmp + ", " + teamN
            tarCell.setValue(tmp)
          }
        }
        else {
          Logger.log("Unable to find " + charN + " character in parseAttenders")
        }
      }
      personTracker = personTracker.offset(1, 0)
    }
    curTeam = curTeam.offset(0, 1)
  }
}

function parseDrops(){
  var lootSheet = BOOK.getSheetByName(LOOT)
  var teamSheet = BOOK.getSheetByName(TEAMS)
  var charSheet = BOOK.getSheetByName(CHAR_POINT)
  var resSheet = BOOK.getSheetByName(RES_IT)
  var limSheet = BOOK.getSheetByName(LIM_IT)
  var unSheet = BOOK.getSheetByName(UN_IT)
  //First open the loot sheet, and read the first entry
  //Then find that character's char_point, their team,
  //and the item's respective page (using the char_point color)
  //Mark the item as dropped for that group (if not already)
  //Remove that item from the char_point for that char
  //Rinse and repeat
  var curName = lootSheet.getRange(2,1).getCell(1,1)
  var curLoot = curName.offset(0, 1)
  var date = teamSheet.getRange("B1").getValue().toString()
  while (curName.getValue().toString().length > 0){
    var tFind = charSheet.createTextFinder(curName.getValue().toString())
    tFind.matchEntireCell(true)
    var tLoc = tFind.findNext()
    if (tLoc == null){
      Logger.log("Unable to find " + curName.getValue().toString() + " character in parseDrops")
      curName = curName.offset(1, 0)
      curLoot = curLoot.offset(1, 0)
      continue
    }
    var tarRange = charSheet.getRange(tLoc.getRow(), 1, 2, 26)
    var teamNames = findTeams(curName.getValue().toString())
    var raidArray = [{}];
    raidArray = teamNames.split(", ");
    var lLoc = tarRange.createTextFinder(curLoot.getValue().toString())
    lLoc.matchEntireCell(true)
    var tarCell = lLoc.findNext()
    var outSheet;
    if (tarCell == null){
      Logger.log("Unable to find " + curLoot.getValue().toString() + " loot item in parseDrops for " + curName.getValue().toString())
      curName = curName.offset(1, 0)
      curLoot = curLoot.offset(1, 0)
      continue
    }
    if (tarCell.getValue().toString().length > 0){
      //First we indicate the loot dropped for that group
      var colorTemp = tarCell.getFontColor()
      if (colorTemp == "#ff0000"){
        outSheet = resSheet
      }
      else if (colorTemp == "#0000ff"){
        outSheet = limSheet
      }
      else {
        outSheet = unSheet
      }
      var outCell = outSheet.createTextFinder(curLoot.getValue().toString()).matchEntireCell(true).findNext()
      var dateCell = outSheet.createTextFinder(date).findNext()
      var raid = outCell.offset(0, 4).getValue().toString()
      for (str in raidArray){
        var loca = raidArray[str].indexOf(raid)
        if (loca > -1){
          var teamName = raidArray[str].charAt(0)
        }
      }
      var entryCell = outSheet.getRange(outCell.getRow(), dateCell.getColumn()).getCell(1,1)
      var temp = entryCell.getValue().toString()
      if (temp.length == 0){
        entryCell.setValue(teamName)
      }
      else {
        entryCell.setValue(temp + ", " + teamName)
      }
      tarCell.setValue("")
    }
    curName = curName.offset(1, 0)
    curLoot = curLoot.offset(1, 0)
  }
}

function findTeams(charName){
  var teamSheet = BOOK.getSheetByName(TEAMS)
  var cFind = teamSheet.createTextFinder(charName)
  var cLoc = cFind.findNext()
  if (cLoc == null){
    Logger.log("Unable to find " + charName + " in findTeams")
    return ""
  }
  var cRange = teamSheet.getRange(2, cLoc.getColumn(), 50, 1)
  var teamName = cRange.getCell(1, 1).getValue().toString()
  cLoc = cFind.findNext()
  if(cLoc != null){
    while(cLoc != null){
      cRange = teamSheet.getRange(2, cLoc.getColumn(), 50, 1)
      teamName = teamName + ", " + cRange.getCell(1, 1).getValue().toString()
      cLoc = cFind.findNext()
    }
  }
  return teamName
}