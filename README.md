# gs-mtg
Google sheets MTG (Magic the Gathering) apps script to fetch card info and prices.  

![Table example](/images/ss1.png)  

The [mtg.gs](mtg.gs) script is intended to be used inside a Google Sheet app script code. To access the apps script code associated with your Google Sheets sheet just click **Tools > Script editor** menu item. More info on Apps script [here](https://developers.google.com/apps-script/guides/sheets). 

To use this script, you will need to "install" an event trigger. Please read [this Google tutorial](https://developers.google.com/apps-script/guides/triggers/installable) to learn how to do it. It will trigger on **edit** events. If you use a simple trigger and try to override the `onEdit(e)` function, it will throw an error stating that you don't have permissions to execute `UrlFetchApp.fetch` function from inside a simple trigger due its restrictions, that's why you need to use an "installable" trigger instead. Please read [this](https://developers.google.com/apps-script/guides/triggers/) for more info on triggers and how to use them.

This script uses the awesome MTG API provided by [magicthegathering.io](https://magicthegathering.io/) Please consider becoming a [patron](https://www.patreon.com/magicthegathering) to support it!
