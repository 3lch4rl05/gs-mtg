//Using Properties service as "cache mechanism" since Cache service 
//only allows maximum of 6hrs before expiring. Since this information
//is very unlikely to change every single day, I decided to use
//properties service instead, which can store properties indefinitely.

var docProps = PropertiesService.getDocumentProperties();
var catGroupsFetched = 0;

var htmlOutput = HtmlService.createHtmlOutputFromFile("loading.html");
htmlOutput.setWidth(300)
htmlOutput.setHeight(150);
var content = htmlOutput.getContent();

function fetchCategoryGroups(limit, offset) {
   
  Logger.log("Fetching groups, limit: " + limit + ", offset: " + offset);
  var response = UrlFetchApp.fetch("https://api.tcgplayer.com/v1.19.0/catalog/categories/1/groups?limit="+limit+"&offset="+offset, options);
  var json = response.getContentText();
  var data = JSON.parse(response);
  for(var i=0; i<data.results.length; i++) {
    var group = data.results[i];
    docProps.setProperty(group.abbreviation, group.name);   
    Logger.log(group.abbreviation + " -> " + docProps.getProperty(group.abbreviation));
  }
  catGroupsFetched+=data.results.length;
  if(catGroupsFetched < data.totalItems) {
    fetchCategoryGroups(limit,catGroupsFetched);
  }
}

function initiateCategoryGroupsFetch(showDialog) {
  
  if(showDialog) {
    showProgressDialog();
  }
  fetchCategoryGroups(50,0);
}

function getGroupByAbr(abr) {
  var group = docProps.getProperty(abr);
  if(!group || group == null) {
    initiateCategoryGroupsFetch(true);
  }
}

function showProgressDialog() {
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Loading sets");
}

function cleanProperties() {
  docProps.deleteAllProperties();
}

function seeProperties() {
  var keys = docProps.getKeys();
  for(var i=0; i<keys.length; i++) {
    Logger.log(docProps.getProperty(keys[i]));
  }
}
