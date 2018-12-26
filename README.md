# gs-mtg
Google sheets MTG (Magic the Gathering) apps script to fetch card info and prices.  

![Table example](/images/ss1.png)  

The [mtg.gs](mtg.gs) script is intended to be used inside a Google Sheet app script code. To access the apps script code associated with your Google Sheets sheet just click **Tools > Script editor** menu item. More info on Apps script [here](https://developers.google.com/apps-script/guides/sheets). 

To use this script, you will need to "install" an event trigger. Please read [this Google tutorial](https://developers.google.com/apps-script/guides/triggers/installable) to learn how to do it. It will trigger on **edit** events. If you use a simple trigger and try to override the `onEdit(e)` function, it will throw an error stating that you don't have permissions to execute `UrlFetchApp.fetch` function from inside a simple trigger due its restrictions, that's why you need to use an "installable" trigger instead. Please read [this](https://developers.google.com/apps-script/guides/triggers/) for more info on triggers and how to use them.

At the top of the script, you can set the column numbers where you want the info to show up. To trigger the API call, you just need to provide the card name and set code in the respective column and that's it! The script will fetch all the info for you. If you pay attention to the code, you will notice that the API call with not be triggered if any of the following are true:

1. The card name length is less than 3 characters.
2. The set code length is less than 3 characters.
3. The multiverse ID column has already a value in it.
4. The active row is below the `STARTING_ROW`.

I have also included some menu items to make it easier to refresh prices and jump to the last row.

![Menu item example](/images/ss2.png)  

This script uses the awesome MTG API provided by [magicthegathering.io](https://magicthegathering.io/) Please consider becoming a [patron](https://www.patreon.com/magicthegathering) to support it!

I'm currently waiting on my dev token application approval from [TCGPlayer](https://www.tcgplayer.com) in order to be able to use their very cool prices API. This will help providing the sell/buy prices for each card in the collection. Will work on it once I get the token. If you are interested in gathering prices as well you should also apply for your own API token. Thanks for checking out my little projects. I hope you find them helpful.
