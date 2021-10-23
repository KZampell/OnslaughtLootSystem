function calcBonus(atten, drops) {
  var bonus = (0.1 * atten) + (0.4 * drops)
  return bonus
}

function fetchAttendance(charN){
  var attSheet = BOOK.getSheetByName(ATT)
  var tFind = attSheet.createTextFinder(charN)
  var loc = tFind.findNext()
  var per = loc.offset(0, 1).getValue()
  if (per == 0){
    return [per, null]
  }
  var attendTuples = []
  var date = attSheet.getRange("F1:F1").getCell(1, 1)
  var marker = loc.offset(0, 5)
  while (date.getRichTextValue().getText().toString().length > 1){
    var tmp = marker.getRichTextValue().getText().toString()
    if (tmp.length > 0){
      if (tmp.length > 2){
        var groups = [{}]
        groups = tmp.split(", ")
        for (ind in groups){
          attendTuples.push(new dropTracker(charN, date.getValue(), groups[ind]))
        }
      }
      else{
        attendTuples.push(new dropTracker(charN, date.getValue(), tmp))
      }
    }
    date = date.offset(0, 1)
    marker = marker.offset(0, 1)
  }
  return [per, attendTuples]
}

function fetchDrops(sName, locCell, attendTuples, dateStart){
  var date = dateStart
  var dropTuples = []
  var cellVal = locCell.offset(0, 5)
  while (date.getRichTextValue().getText().toString().length > 1){
    var temp = cellVal.getRichTextValue().getText().toString()
    if (temp.length > 0){
      dropTuples.push(new dropTracker(sName, date.getValue(), temp))
    }
    date = date.offset(0, 1)
    cellVal = cellVal.offset(0,1)
  }
  var count = 0
  for (var dT in dropTuples){
    var sDate = dropTuples[dT].date
    var sGrp = dropTuples[dT].group
    for (var aT in attendTuples){
      var cDate = attendTuples[aT].date
      if (sDate == cDate){
        var cGrp = attendTuples[aT].group
        //Logger.log("char %s, item %s", cGrp, sGrp)
        if (sGrp.indexOf(cGrp) > -1){
          count++
          break
        }
      }
    }
  }
  return count
}

function countTies(iName, cName){
  var tieSheet = BOOK.getSheetByName(PASS)
  var cellA = tieSheet.getRange("A2")
  var cellB = cellA.offset(0,1)
  var count = 0
  while (cellA.getRichTextValue().getText().toString().length > 1){
    var txtA = cellA.getValue().toString()
    if (txtA == cName){
      var txtB = cellB.getValue().toString()
      if (txtB == iName){
        count++
      }
    }
    cellA = cellA.offset(1,0)
    cellB = cellB.offset(1,0)
  }
  return count
}

function interpretPasses(){
  var tieSheet = BOOK.getSheetByName(PASS)
  var lootSheet = BOOK.getSheetByName(LOOT)
  var teamSheet = BOOK.getSheetByName(TEAMS)
  var p1Sheet = BOOK.getSheetByName(P1)
  var p2Sheet = BOOK.getSheetByName(P2)
  var p3Sheet = BOOK.getSheetByName(P3)
  var curLoot = lootSheet.getRange("A2:C2")
  var curPass = tieSheet.getRange("A2:C2")
  while (curPass.getValue().toString().length > 1){
    curPass = curPass.offset(1,0)
  }
  while(curLoot.getValue().toString().length > 1){
    if (curLoot.getFontWeight() == "bold"){
      curLoot = curLoot.offset(1,0)
      continue
    }
    var tarItem = curLoot.getCell(1,2).getValue().toString()
    var tarName = curLoot.getCell(1,1).getValue().toString()
    var tarDate = curLoot.getCell(1,3).getValue().toString()
    var tFind = p1Sheet.createTextFinder(tarItem)
    tFind.matchEntireCell(true)
    var loc = tFind.findNext()
    var activeSheet = p1Sheet
    if (loc == null){
      tFind = p2Sheet.createTextFinder(tarItem)
      tFind.matchEntireCell(true)
      loc = tFind.findNext()
      activeSheet = p2Sheet
      if (loc == null){
        tFind.matchEntireCell(false)
        loc = tFind.findNext()
      }
      if (loc == null){
        tFind = p3Sheet.createTextFinder(tarItem)
        tFind.matchEntireCell(true)
        loc = tFind.findNext()
        activeSheet = p3Sheet
        if (loc == null){
        tFind.matchEntireCell(false)
        loc = tFind.findNext()
        if (loc == null){
          Logger.log("I could not find the object " + tarItem)
          curLoot = curLoot.offset(1,0)
          continue
        }
        }
      }
    }
    //Found it
    Logger.log("I found the object " + tarItem)
    var tRow = loc.getRow()
    var cRow = activeSheet.getRange(tRow, 3, 1, 23)
    var cFind = cRow.createTextFinder(tarName)
    var cCell = cFind.findNext()
    if (cCell != null){
      var tCol = loc.getColumn()
      var teamCol = teamSheet.getRange(1, tCol, 50, 1)
      var attended = []
      for (var i = 1; i < 50; i++){
        if (teamCol.getCell(i, 1).getBackground() == "#00ff00"){
          attended.push(teamCol.getCell(i, 1).getValue().toString())
        }
      }
      var names = [{}]
      names = cCell.getValue().toString().split(", ")
      Logger.log(names)
      var score
      var hitTar = false
      for(var name in names){
        var sc, nm
        [sc, nm] = names[name].split(" : ")
        if (name == 0){
          score = sc
        }
        if (nm == tarName){
          hitTar = true
          score = sc
          continue
        }
        else if (hitTar & sc < score){
          break
        }
        else if (attended.indexOf(nm) >= 0) {
          curPass.getCell(1,1).setValue(nm)
          curPass.getCell(1,2).setValue(tarItem)
          curPass.getCell(1,3).setValue(tarDate)
          curPass = curPass.offset(1,0)
          Logger.log(tarItem + " marked for " + nm)
        }
      }
    } else {
      Logger.log("Item " + tarItem + " given to " + tarName + " was not a priority item.")
    }
    curLoot.setFontWeight("bold")
    curLoot = curLoot.offset(1,0)
  }
}