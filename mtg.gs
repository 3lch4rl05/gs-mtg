var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName("Coleccion");

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

function onOpen(e) {
  
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Utils")
    .addItem("Jump to last", "jumpToLast")
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Refresh')
      .addItem("Refresh Prices", "refreshPrices")
      .addItem("Refresh All", "refreshAll"))
    .addToUi();
}

function fetchCard(e) {

  var range = e.range;
  var a1 = range.getA1Notation();
  var col = range.getColumn();
  var row = range.getRow();
  var val = range.getValue();
  var cardInfo;
  
  var celdaNombre = sheet.getRange(row, NAME_COL);
  var celdaSerie = sheet.getRange(row, SET_COL);
  
  var valNombre = celdaNombre.getValue();
  var valSerie = celdaSerie.getValue();
  
  if(sheet.getRange(row,MULTIVERSEID_COL).getValue().length == 0) {
    if(col == NAME_COL || col == SET_COL) {
      if(row >= 3) {
        if(val.length >= 3) {
          if(col == NAME_COL) {
            if(valSerie.length >= 3) {
              cardInfo = getCardInfo(val,valSerie);
            }
          } else if(col == SET_COL) {
            if(valNombre.length >= 3) {
              cardInfo = getCardInfo(valNombre, val);
            }
          }
        }
      }
    }
  }
  
  if(cardInfo.cards.length>0) {
  
    var card = cardInfo.cards[0];
    
    //Multiverse ID
    sheet.getRange(row, MULTIVERSEID_COL).setValue(card.multiverseid);
    
    //Set name
    sheet.getRange(row, SETNAME_COL).setValue(card.setName);
    
    //Set icon
    sheet.getRange(row, SETICON_COL).setValue("=IMAGE(\"http://gatherer.wizards.com/Handlers/Image.ashx?type=symbol&set=" + card.set + "&size=medium\")");
    
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
    if(card.colors) {
      var colorsQty = card.colors.length;
      for(var i=0; i<colorsQty; i++) {
        colorCS = colorCS + card.colors[i];
        if(i<(colorsQty-1)) {
          colorCS = colorCS + ", ";
        }
      }
    }
    sheet.getRange(row, COLOR_COL).setValue(colorCS);
    
    sheet.setRowHeight(row, 177);
    sheet.setColumnWidth(IMAGE_COL, 127);
    
    celdaNombre.setFontColor("#000000");
    celdaSerie.setFontColor("#000000");
    
  } else {
  
    celdaNombre.setFontColor("#FF0000");
    celdaSerie.setFontColor("#FF0000");
  }
}

function getCardInfo(name, set) {

  var response = UrlFetchApp.fetch("https://api.magicthegathering.io/v1/cards?name="+name+"&set="+set);
  var json = response.getContentText();
  var data = JSON.parse(response);
  return data;
}

function refreshRow(row) {
  
  Logger.log("Refreshing row: " + row);
  var e = {
    range : sheet.getRange(row, NAME_COL)
  };
  
  sheet.getRange(row,MULTIVERSEID_COL).setValue("");
  fetchCard(e);
}

function refreshPrices() {
  
}

function refreshAll() {

  Logger.log("Refreshing all cards info...");
  for(var i=STARTING_ROW; i<=sheet.getLastRow(); i++) {
    refreshRow(i);
  }
}

function jumpToLast() {
   sheet.setActiveRange(sheet.getRange(sheet.getLastRow(),3));
}  
