var doneBool

function timeFunction() {
  //Check the Info page for Flags, set flags, and proceed as needed.
  startTime = new Date().getTime()
  var currentPage = BOOK.getSheetByName("INFO")
  var phaseCell = currentPage.getRange("G1").getCell(1,1)
  var curPhase = phaseCell.getValue().toString()
  doneBool = false
  if (curPhase == "0"){
    //Cleanup the sheets
    wipeLists()
    emptyStorage()
    phaseCell.setValue("1")
  } else if (curPhase == "1"){
    //Process the data
    updateItemValues()
    if (doneBool){
      phaseCell.setValue("2")
    }
  } else if (curPhase == "2"){
    //Sort Data
    sortData()
    phaseCell.setValue("3")
  } else if (curPhase == "3"){
    //Process the Data P1
    //fillTuples()
    updateSheet(P1)
    if (doneBool){
      phaseCell.setValue("4")
    }
  } else if (curPhase == "4"){
    //Process the Data P2
    fillTuples()
    updateSheet(P2)
    if (doneBool){
      phaseCell.setValue("5")
    }
  } else if (curPhase == "5"){
    //Process the Data P3
    fillTuples()
    updateSheet(P3)
    if (doneBool){
      phaseCell.setValue("6")
    }
  } else if (curPhase == "6"){
    //Cleanup the Data
    cleanLists()
    phaseCell.setValue("ReadyToRaid")
  } else if (curPhase == "7"){
    // Clean Up Pasted List
    formatLootSheet()
    phaseCell.setValue("8")
  } else if (curPhase == "8"){
    //Interpret Passes
    interpretPasses()
    phaseCell.setValue("9")
  } else if (curPhase == "9"){
    //Parse attendance
    parseAttenders()
    phaseCell.setValue("10")
  } else if (curPhase == "10"){
    //Post Raid Function
    parseDrops()
    phaseCell.setValue("AwaitingNewTeams")
  }
}