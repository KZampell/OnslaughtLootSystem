function sortData(){
  var storage = BOOK.getSheetByName(DAT)
  var activeItem = storage.getRange("A1:E1")
  while (!activeItem.getCell(1,1).getRichTextValue().getTextStyle().isItalic()){
    var itemCell = activeItem.getCell(1,1).getValue().toString()
    var scoreCell = activeItem.getCell(1, 2).getValue()
    var nameCell = activeItem.getCell(1,3).getValue().toString()
    var classCell = activeItem.getCell(1,4).getValue()
    var teamCell = activeItem.getCell(1,5).getValue().toString()
    TUPLES.push(new tEntry(itemCell, scoreCell, nameCell, classCell, teamCell))
    activeItem = activeItem.offset(1,0)
  }
  activeItem = activeItem.offset(1,0)
  
  TUPLES.sort(function(x,y) {
    var xp = x.item;
    var yp = y.item;
    var xq = x.score;
    var yq = y.score;
    return xp < yp ? -1 : xp > yp ? 1 : xq > yq ? -1 : xq < yq ? 1 : 0;
    // return 0 if equal, 1 if > and -1 if <
  })

  storage.clear()

  activeItem = storage.getRange("A1:E1")

  for(var c = 0; c < TUPLES.length; c++){
    var itemCell = activeItem.getCell(1,1)
    var scoreCell = activeItem.getCell(1, 2)
    var nameCell = activeItem.getCell(1,3)
    var classCell = activeItem.getCell(1,4)
    var teamCell = activeItem.getCell(1,5)
    var en = TUPLES[c]
    itemCell.setValue(en.item)
    scoreCell.setValue(en.score.toString())
    nameCell.setValue(en.char)
    classCell.setValue(en.charClass)
    teamCell.setValue(en.team)

    activeItem = activeItem.offset(1, 0)
  }

  activeItem.getCell(1,1).setValue("End of Line")
  activeItem.getCell(1,1).setFontStyle("italic")
  
}

function fillTuples(){
  var storage = BOOK.getSheetByName(DAT)
  var activeItem = storage.getRange("A1:E1")
  while (!activeItem.getCell(1,1).getRichTextValue().getTextStyle().isItalic()){
    var itemCell = activeItem.getCell(1,1).getValue().toString()
    var scoreCell = activeItem.getCell(1, 2).getValue()
    var nameCell = activeItem.getCell(1,3).getValue().toString()
    var classCell = activeItem.getCell(1,4).getValue()
    var teamCell = activeItem.getCell(1,5).getValue().toString()
    TUPLES.push(new tEntry(itemCell, scoreCell, nameCell, classCell, teamCell))
    activeItem = activeItem.offset(1,0)
  }
  
}

function formatLootSheet(){
  var shit = BOOK.getSheetByName(LOOT)
  var activeCell = shit.getRange("A2:C2")
  while(activeCell.getCell(1,1).getValue().toString().length > 1){
    var name, server
    [name, server] = activeCell.getCell(1,1).getValue().toString().split("-")
    activeCell.getCell(1,1).setValue(name)
    var item = activeCell.getCell(1, 2).getValue().toString()
    activeCell.getCell(1,2).setValue(item.substr(1, item.length-2))
    var date = activeCell.getCell(1,3).getValue().toString()
    activeCell.getCell(1,3).setValue(date.substr(0, date.length-3))
    activeCell = activeCell.offset(1, 0)
  }
}