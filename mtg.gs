var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName("Collection");
var ui = SpreadsheetApp.getUi();

var MINIMUM_LENGTH = 3;
var STARTING_ROW = 3;
var MULTIVERSEID_COL = 2;
var NAME_COL = 5;
var SET_COL = 6;
var SETNAME_COL = 7;
var SETICON_COL = 8;
var NUMBER_COL = 9;
var TYPE_COL = 10;
var RARITY_COL = 11;
var IMAGE_COL = 12;
var COLOR_COL = 13;

var IMG_HEIGHT = 152;
var IMG_WIDTH = 106;

function fetchCardData(e) {

  var range = e.range;
  var a1 = range.getA1Notation();
  var col = range.getColumn();
  var row = range.getRow();
  var fireCall = false;
  
  Logger.log("Activated on edit for range: " + a1);
  var name = sheet.getRange(row, NAME_COL);
  var set = sheet.getRange(row, SET_COL);
  var number = sheet.getRange(row, NUMBER_COL);
  
  if(sheet.getRange(row,MULTIVERSEID_COL).getValue().length == 0) {
    if(row >= STARTING_ROW) {
      if(col == NAME_COL || col == SET_COL || col == NUMBER_COL) {
        if(number.getValue() > 0
          && set.getValue().length >= MINIMUM_LENGTH) {
          Logger.log("Call fired by card number!");
          fireCall = true;
        }
        else if(name.getValue().length >= MINIMUM_LENGTH 
          && set.getValue().length >= MINIMUM_LENGTH) {
          Logger.log("Call fired");
          fireCall = true;
        }
      }
    }
  }
  
  if(fireCall == true) {
  
    var cardInfo = getCardInfo(name.getValue(), set.getValue(), number.getValue());    
    if(cardInfo && cardInfo.cards.length>0) {
    
      var card = cardInfo.cards[0];
      var rarityLetter = "C";
      if(card.rarity != "Basic Land") {
        rarityLetter = card.rarity.charAt(0); 
      }
      
      //Multiverse ID
      sheet.getRange(row, MULTIVERSEID_COL).setValue(card.multiverseid);
      
      //Card name
      if(sheet.getRange(row, NAME_COL).getValue().length == 0) {
        sheet.getRange(row, NAME_COL).setValue(card.name);
      }
      
      //Set name
      sheet.getRange(row, SETNAME_COL).setValue(card.setName);
      
      //Set icon
      sheet.getRange(row, SETICON_COL).setValue("=IMAGE(\"http://gatherer.wizards.com/Handlers/Image.ashx?type=symbol&set=" + card.set + "&size=medium&rarity=" + rarityLetter + "\")");
      
      //Card number
      sheet.getRange(row, NUMBER_COL).setValue(card.number);
      
      //Type
      sheet.getRange(row, TYPE_COL).setValue(card.originalType);
      
      //Rarity
      sheet.getRange(row, RARITY_COL).setValue(card.rarity);
      
      //Image
      sheet.getRange(row, IMAGE_COL).setValue("=IMAGE(\"" + card.imageUrl + "\")");
      
      //Color
      var colorCS = "";
      if(!card.originalType.indexOf("Basic Land") == 0 && card.colors) {
        var colorsQty = card.colors.length;
        for(var i=0; i<colorsQty; i++) {
          colorCS = colorCS + card.colors[i];
          if(i<(colorsQty-1)) {
            colorCS = colorCS + ", ";
          }
        }
      } else {
        colorCS = card.watermark;
      }
      
      sheet.getRange(row, COLOR_COL).setValue(colorCS);
      sheet.setRowHeight(row, IMG_HEIGHT);
      sheet.setColumnWidth(IMAGE_COL, IMG_WIDTH);
      
      name.setFontColor("#000000");
      set.setFontColor("#000000");
      
    } else {
    
      name.setFontColor("#FF0000");
      set.setFontColor("#FF0000");
    }
  }
}

function getCardInfo(name, set, number) {

  var response;
  if(!number) {
    Logger.log("Fetching card info for: [" + name + "," + set + "]");
    response = UrlFetchApp.fetch("https://api.magicthegathering.io/v1/cards?name="+name+"&set="+set);
  } else {
    Logger.log("Fetching card info for: [" + set + "," + number + "]");
    response = UrlFetchApp.fetch("https://api.magicthegathering.io/v1/cards?number="+number+"&set="+set);
  }
  
  var json = response.getContentText();
  var data = JSON.parse(response);
  return data;
}

function refreshPrices() {
  //waiting on my token approval...
}

function refreshRow(row) {
  
  Logger.log("Refreshing row: " + row);
  var e = {
    range : sheet.getRange(row, NAME_COL)
  };
  
  sheet.getRange(row,MULTIVERSEID_COL).setValue("");
  fetchCardData(e);
}

function refreshFromRow(row) {

  Logger.log("Refreshing cards info from row " + row + "...");
  for(var i=row; i<=sheet.getLastRow(); i++) {
    refreshRow(i);
  }
}

function refreshAll() {

  Logger.log("Refreshing all cards info...");
  refreshFromRow(STARTING_ROW);
}

function jumpToLast() {
   sheet.setActiveRange(sheet.getRange(sheet.getLastRow(),3));
}

function onOpen(e) {
  
  ui.createMenu("Utils")
    .addItem("Jump to last", "jumpToLast")
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Refresh')
      .addItem("Refresh Row", "askForRowToRefresh")
      .addItem("Refresh From Row", "askForRowRefreshFrom")
      .addItem("Refresh Prices Only", "refreshPrices")
      .addItem("Refresh All", "refreshAll"))
    .addToUi();
}

function askForRow() {
  var response = ui.prompt("Row number:");
  return response.getResponseText();
  
}

function askForRowToRefresh() {
  refreshRow(askForRow());
}

function askForRowRefreshFrom() {
  refreshFromRow(askForRow());
}
