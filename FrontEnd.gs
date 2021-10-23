function weeklyChars(){
  startTime = new Date().getTime()
  BOOK.toast('Calculating Item Values...')
  updateItemValues()
  endTime = new Date().getTime()
  var t = (endTime-startTime)/1000
  Logger.log("Process took " + t + " seconds")
}

function weeklyUpdate1() {
  startTime = new Date().getTime()
  fillTuples()
  BOOK.toast('Populating P1')
  updateSheet(P1)
  endTime = new Date().getTime()
  var t = (endTime-startTime)/1000
  Logger.log("Process took " + t + " seconds")
}

function weeklyUpdate2() {
  startTime = new Date().getTime()
  fillTuples()
  BOOK.toast('Populating P2')
  updateSheet(P2)
  endTime = new Date().getTime()
  var t = (endTime-startTime)/1000
  Logger.log("Process took " + t + " seconds")
}

function weeklyUpdate3() {
  startTime = new Date().getTime()
  fillTuples()
  BOOK.toast('Populating P3')
  updateSheet(P3)
  endTime = new Date().getTime()
  var t = (endTime-startTime)/1000
  Logger.log("Process took " + t + " seconds")
}

function wipeLists(){
  // This function takes approximately 2 minutes to execute
  wipeSlate(P1)
  wipeSlate(P2)
  wipeSlate(P3)
}

function postRaidFunction(){
  startTime = new Date().getTime()
  BOOK.toast('Updating Drop Record and Character Loot Lists')
  parseDrops()
  endTime = new Date().getTime()
  var t = (endTime-startTime)/1000
  Logger.log("Process took " + t + " seconds")
}

function verifyLists(){
  colorLists()
  parsePicks()
}

function emptyStorage(){
  var storage = BOOK.getSheetByName(DAT)
  storage.clear()
}

function cleanLists(){
  var charSheet = BOOK.getSheetByName(CHAR_POINT)
  var rng = charSheet.getRange("A:A")
  rng.setFontWeight('normal')
  cleanupPage(P1)
  cleanupPage(P2)
  cleanupPage(P3)
}

function cleanupPage(page){
  var mainSheet = BOOK.getSheetByName(page)
  var curIter = mainSheet.getRange("A2").getCell(1,1)
  while(curIter.getFontStyle() != 'italic'){
    if(curIter.getBackground() == "#ffffff"){
      curIter.setFontWeight('normal')
    }
    curIter = curIter.offset(1, 0)
  }
}

function onOpen(){
  var user = Session.getActiveUser().getEmail()
  Logger.log(user)
  var ui = SpreadsheetApp.getUi()
  if(adminList.indexOf(user) > -1){   
      ui.createMenu('Admin Options')
        .addItem('Color Code & Verify Lists', 'verifyLists')
        .addItem('Fix the loot sheet after pasting', 'formatLootSheet')
        .addItem('Populate pass/tie list', 'interpretPasses')
        .addItem('Run to parse attendance record', 'parseAttenders')
        .addItem('Run After All Raid Drops Input', 'postRaidFunction')
        .addItem('Clear Off Drop Tables', 'wipeLists')  
        .addItem('Parse Data for Drop Tables', 'weeklyChars')
        .addItem('Sort parsed data', 'sortData')
        .addItem('Update P1 Drop Tables', 'weeklyUpdate1')
        .addItem('Update P2 Drop Tables', 'weeklyUpdate2')
        .addItem('Update P3 Drop Tables', 'weeklyUpdate3')
        .addItem('Purge Data Storage', 'emptyStorage')
        .addItem('Cleanup the Mess', 'cleanLists')
        .addToUi()
  }
}

function adminButton(){
  var user = Session.getActiveUser().getEmail()
  Logger.log(user)
  if(adminList.indexOf(user) > -1){
    onOpen()
  }
}