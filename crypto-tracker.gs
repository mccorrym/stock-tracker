/*
Not needed any longer when using the Coinbase API
This was a hack to trick =CRYPTOFINANCE into updating the sheet each time it changed.
function populateRandomTime() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Overview").getRange("K31").setValue(Utilities.formatDate(new Date(), "GMT", "yyyyMMddHHmmss"));
}
*/
function GET_BTC_PRICE() {
  var response = UrlFetchApp.fetch("https://api.coinbase.com/v2/prices/BTC-USD/buy");
  var w = JSON.parse(response);
  // The amount returned by the API includes a 1% Coinbase fee
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Overview").getRange("H30").setValue(parseFloat(w["data"]["amount"]) - (parseFloat(w["data"]["amount"]) * .01));
}
