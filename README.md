# gs-mtg
Google sheets MTG (Magic the Gathering) apps script to fetch card info and prices.  

![Table example](/images/ss1.png)  

The [mtg.gs](mtg.gs) script is intended to be used inside a Google Sheet app script code. To access the apps script code associated with your Google Sheets sheet just click **Tools > Script editor** menu item. More info on Apps script [here](https://developers.google.com/apps-script/guides/sheets). 

To use this script, you will need to "install" an event trigger. Please read [this Google tutorial](https://developers.google.com/apps-script/guides/triggers/installable) to learn how to do it. It will trigger on **edit** events. If you use a simple trigger and try to override the `onEdit(e)` function, it will throw an error stating that you don't have permissions to execute `UrlFetchApp.fetch` function from inside a simple trigger due its restrictions, that's why you need to use an "installable" trigger instead. Please read [this](https://developers.google.com/apps-script/guides/triggers/) for more info on triggers and how to use them.

At the top of the script, you can set the column numbers where you want the info to show up. To trigger the API call, you just need to provide the card name and set code, or the set code and card number in the respective columns and that's it! The script will fetch the rest of the info for you. If you pay attention to the code, you will notice that the API call will be triggered only if:

__Fetching by card name and set code__  

1. The card name length is equal or longer than 3 characters.
2. The set code length is equal or longer than 3 characters.

__Fetching by card number and set code__

1. The card number is greater than 0 (zero).
2. The set code length is equal or longer than 3 characters.

__For all scenarios__

* The multiverse ID column is empty.
* The active row is greater than the `STARTING_ROW`.

In case you provide invalid values for the card name, set code or card number, the script will let you know by setting the font color to red like shown in the image below. You just need to update the colum with the correct information and it will update the rest for you.

<img src="/images/ss3.png" width="500px">

I have also included some menu items to make it easier to refresh info and jump to the last row.

<img src="/images/ss2.png" width="500px">

This script uses the awesome MTG API provided by [magicthegathering.io](https://magicthegathering.io/) Please consider becoming a [patron](https://www.patreon.com/magicthegathering) to support it!

Also, I recently started using [Scryfall](https://scryfall.com) to improve a little bit how the script works, specially to fetch a higher quality image and also, very important, to finally be able to fetch prices! Please also consider [donating](https://scryfall.com/donate) so they can continue with the really cool work!
