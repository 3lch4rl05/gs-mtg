var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName("Collection");
var ui = SpreadsheetApp.getUi();

var MINIMUM_LENGTH = 3;
var STARTING_ROW = 3;
var MULTIVERSEID_COL = 2;
var FOIL_COL = 4;
var NAME_COL = 5;
var SET_COL = 6;
var SETNAME_COL = 7;
var SETICON_COL = 8;
var NUMBER_COL = 9;
var TYPE_COL = 10;
var RARITY_COL = 11;
var IMAGE_COL = 12;
var COLOR_COL = 13;
var PRICE_COL = 14;

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
      var isFoil = sheet.getRange(row, FOIL_COL).getValue() == "Si";
      var price = getCardBuyPrice(card.name, getGroupByAbr(set.getValue()), isFoil);
      var rarityLetter = "C";
      if(card.rarity != "Basic Land") {
        rarityLetter = card.rarity.charAt(0); 
      }
      
      //Multiverse ID
      sheet.getRange(row, MULTIVERSEID_COL).setValue(card.multiverseid);
      
      //Card name
      sheet.getRange(row, NAME_COL).setValue("=HYPERLINK(\"https://img.scryfall.com/cards/border_crop/en/" 
            + card.set.toLowerCase() + "/" + card.number + ".jpg\",\"" + card.name + "\")");
            
      //Text
      if(card.rarity != "Basic Land") {
        sheet.getRange(row, NAME_COL).setNote(card.text);
      }
      
      //Set name
      sheet.getRange(row, SETNAME_COL).setValue(card.setName);
      
      //Set icon
      sheet.getRange(row, SETICON_COL).setValue("=IMAGE(\"http://gatherer.wizards.com/Handlers/Image.ashx?type=symbol&set=" 
          + card.set + "&size=medium&rarity=" + rarityLetter + "\")");
      
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
      } else if(card.rarity == "Basic Land"
          && card.watermark != "Colorless") {
        colorCS = card.watermark;
      }
      
      sheet.getRange(row, COLOR_COL).setValue(colorCS);
      
      //Price
      sheet.getRange(row, PRICE_COL).setValue(price);
      
      sheet.setRowHeight(row, IMG_HEIGHT);
      sheet.setColumnWidth(IMAGE_COL, IMG_WIDTH);
      
      number.setFontColor("#000000");
      name.setFontColor("#000000");
      set.setFontColor("#000000");
      
    } else {
    
      number.setFontColor("#FF0000");
      name.setFontColor("#FF0000");
      set.setFontColor("#FF0000");
    }
  }
}

function getCardInfo(name, set, number) {

  var response;
  var params;
  
  if(!number) {
    Logger.log("Fetching card info for: [" + name + "," + set + "]");
    params = "name="+name+"&set="+set;
  } else {
    Logger.log("Fetching card info for: [" + set + "," + number + "]");
    params = "number="+number+"&set="+set;
  }
  
  response = UrlFetchApp.fetch("https://api.magicthegathering.io/v1/cards?" + params);
  var json = response.getContentText();
  var data = JSON.parse(response);
  return data;
}

function getCardProdId(cardName, setName) {
  
  var search_filters = {
	"filters": [
      {
        "name":"ProductName",
		"values": [
          cardName
		]
      },
      {
		"name":"SetName",
		"values": [
          setName
		]
      }
	]
  }
  
  options_search.payload = JSON.stringify(search_filters);
  Logger.log("Filters: " + options_search.payload);
  var response = UrlFetchApp.fetch("https://api.tcgplayer.com/v1.19.0/catalog/categories/1/search/", options_search);
  if(response.getResponseCode() == 200) {
    var json = response.getContentText();
    var data = JSON.parse(response);
    Logger.log(data);
    var prodId = data.results[0];
    return prodId;
  } else {
    Logger.log(response)
    return null;
  }
}

function getCardSKU(prodId, printing) {

  var response = UrlFetchApp.fetch("https://api.tcgplayer.com/v1.19.0/catalog/products/" + prodId + "/skus", options);
  var json = response.getContentText();
  var data = JSON.parse(response);
  for(var i=0; i<data.results.length; i++) {
    var sku = data.results[i];    
    if(sku.languageId == ENGLISH
      && sku.printingId == printing
      && sku.conditionId == NEAR_MINT) {
        Logger.log(sku.skuId);
        return sku.skuId;
    }
  }  
}

function getPriceBySKU(skuId) {

  var response = UrlFetchApp.fetch("http://api.tcgplayer.com/v1.19.0/pricing/buy/sku/" + skuId, options);
  var json = response.getContentText();
  var data = JSON.parse(response);
  return data.results[0].prices.market;
}

function getCardBuyPrice(cardName, setName, isFoil) {
  
  var prodId = getCardProdId(cardName, setName);
  Logger.log("ProdID found: " + prodId);
  if(prodId != null) {
  
    var printing = NORMAL_PRINT;
    
    if(isFoil) {
      printing = FOIL_PRINT;
    }
    
    var skuId = getCardSKU(prodId, printing);
    Logger.log("SKU for: " + prodId + " and " + printing + ": " + skuId);
    var buyPrice = getPriceBySKU(skuId);
    Logger.log("Price: " + buyPrice);
    return buyPrice;
  } else {
    return "N/A";
  }
}
 
function refreshBuyPrices(row, lastRow) {

  if(!lastRow) {
    lastRow = sheet.getLastRow();
  }

  for(var i=row; i<=lastRow; i++) {
    var name = sheet.getRange(i, NAME_COL).getValue();
    var setName = getGroupByAbr(sheet.getRange(i, SET_COL).getValue());
    var isFoil = sheet.getRange(i, FOIL_COL).getValue() == "Si";
    var newPrice = getCardBuyPrice(name, setName, isFoil);
    sheet.getRange(i, PRICE_COL).setValue(newPrice);
  }
}

function refreshRow(row) {
  
  Logger.log("Refreshing row: " + row);
  var e = {
    range : sheet.getRange(row, NAME_COL)
  };
  
  sheet.getRange(row,MULTIVERSEID_COL).setValue("");
  fetchCardData(e);
}

function refreshFromRow(row, lastRow) {

  if(!lastRow) {
    lastRow = sheet.getLastRow();
  }
  
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
      .addItem("Refresh Prices Only", "askForRowToRefreshPrices")
      .addItem("Refresh sets cache", "fetchGroupsWithModal")
      .addItem("Refresh All", "refreshAll"))
    .addToUi();
}

function askForStartRow() {
  var response = ui.prompt("Start from row number:");
  return response.getResponseText();
}

function askForEndRow() {
  var response = ui.prompt("End row number:");
  return response.getResponseText();
}

function askForRowToRefresh() {
  refreshRow(askForStartRow());
}

function askForRowRefreshFrom() {
  refreshFromRow(askForStartRow(), askForEndRow());
}

function askForRowToRefreshPrices() {
  refreshBuyPrices(askForStartRow(), askForEndRow());
}

function fetchGroupsWithModal() {
  initiateCategoryGroupsFetch(true);
}
