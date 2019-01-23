var tickers = ["SCHB", "SCHA", "SCHX", "SCHF", "SCHE", "SCHD", "TFI", "C", "BRK.B", "F", "AAPL", "APTV", "DLPH"];
var options = {
  headers: {
    "Cache-Control": "max-age=0"
  }
};

function FIND_TODAYS_CELL(sheet_name) {
  var values = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange("A:A").getValues();
  var current_date = new Date();
  current_date.setHours(0,0,0,0)
  var column_count = 0;
  while (values[column_count][0] != "") {
    var cell_date = new Date(values[column_count][0]);
    if (current_date.getTime() == cell_date.getTime()) {
      return (column_count + 1)
    }
    column_count++;
  }
  return false;
}
function GET_REALTIME_PRICING() {
  var current_date = new Date();
  if (current_date.getDay() > 0 && current_date.getDay() < 6) {
    var str = "";
    for (i=0; i<tickers.length; i++) {
      str += tickers[i]+",";
    }
    str = str.substr(0, (str.length-1));
    var response = UrlFetchApp.fetch("https://www.alphavantage.co/query?function=BATCH_QUOTES_US&symbols="+str+"&apikey="+PropertiesService.getScriptProperties().getProperty("api_key"), options);
    var json = JSON.parse(response);
    for (var key in json["Stock Batch Quotes"]) {
      if (!json["Stock Batch Quotes"].hasOwnProperty(key)) continue;
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(json["Stock Batch Quotes"][key]["1. symbol"]).getRange("F2").setValue(json["Stock Batch Quotes"][key]["5. price"]);
      if (current_date.getHours() > 15) {
        // For some reason the TIME_SERIES_DAILY quotes are not accurate after-hours.
        // We need to correct the after-hours pricing with the BATCH_QUOTES_US results.
        // This will only be done during the same trading day as the closing date, after the cell has been updated at least once with the TIME_SERIES_DAILY result.
        var cell_number = FIND_TODAYS_CELL(json["Stock Batch Quotes"][key]["1. symbol"]);
        if (cell_number !== false) {
          SpreadsheetApp.getActiveSpreadsheet().getSheetByName(json["Stock Batch Quotes"][key]["1. symbol"]).getRange(cell_number, 2).setValue(json["Stock Batch Quotes"][key]["5. price"]);
        }
      }
    }
  }
}
function GET_HISTORICAL_PRICING() {
  var current_date = new Date();
  var ticker_index = parseInt(PropertiesService.getScriptProperties().getProperty("ticker_index"));
  
  if (!isNaN(ticker_index)) {
    for (i=ticker_index; i<(ticker_index+3); i++) {
      var response = UrlFetchApp.fetch("https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&outputsize=full&symbol="+tickers[i]+"&apikey="+PropertiesService.getScriptProperties().getProperty("api_key"), options);
      var json = JSON.parse(response);
      var prices = [];
      for (var key in json["Time Series (Daily)"]) {
        if (!json["Time Series (Daily)"].hasOwnProperty(key)) continue;
        
        var close_date = new Date(key);
        
        if (close_date.getFullYear() == current_date.getFullYear()) {
          prices.unshift({
            "date": key,
            "close": json["Time Series (Daily)"][key]["4. close"]
          });
        } else {
          break;
        }
      }
      
      if (prices.length == 0) {
        TRY_AGAIN(i);
        return false;
      }
      
      var quote_date = new Date(prices[(prices.length-1)]["date"]);
      
      // Add one day to the quote date (reports as 1 day in the past for some reason)
      var quote_current_date = new Date(quote_date.getTime() + 86400000);
      
      if (quote_current_date.getDate() != current_date.getDate()) {
        TRY_AGAIN(i);
        return false;
      } else {
        PropertiesService.getScriptProperties().deleteProperty("ticker_tries");
      }

      for (j=0; j<prices.length; j++) {
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tickers[i]).getRange("A"+(j+2)).setValue(prices[j]["date"]);
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tickers[i]).getRange("B"+(j+2)).setValue(prices[j]["close"]);
      }
      if ((i+1) == tickers.length) {
        PropertiesService.getScriptProperties().deleteProperty("ticker_index");
        return false;
      }
    }
    PropertiesService.getScriptProperties().setProperty("ticker_index", (ticker_index+3));
  } else {
    // Check at 10AM on weekdays to see whether the market is open. This is done to avoid updating any sheets with closing prices erroneously in the event of a market holiday.
    if (current_date.getDay() > 0 && current_date.getDay() < 6 && current_date.getHours() == 10 && current_date.getMinutes() == 0) {
      // The Alphavantage and/or IPX APIs don't currently offer this capability. But Apple does! (for now?)
      var url = "http://wu-quotes.apple.com/dgw?imei=555&apptype=finance";
      var options = {
        "method": "post",
        "headers": {
          "Content-Type": "text/xml"
        },
        "payload": "<?xml version='1.0' encoding='utf−8'?><request devtype='Apple_OSX' deployver='APPLE_DASHBOARD_1_0' app='YGoAppleStocksWidget' appver='unknown' api='finance' apiver='1.0.1' acknotification='0000'><query id='0' timestamp='`date +%s000`' type='getquotes'><list><symbol>SCHD</symbol></list></query></request>"
      };
      var response = UrlFetchApp.fetch(url, options);
    
      var document = XmlService.parse(response);
      var root = document.getRootElement();
      
      try {
        // The market_status value is 1 during market hours.
        var market_status = root.getChild("result").getChild("list").getChild("quote").getChild("status").getText();
      
        if (market_status != "1") {
          PropertiesService.getScriptProperties().setProperty("market_open", false);
        } else {
          PropertiesService.getScriptProperties().setProperty("market_open", true);
        }
      } catch(e) {
        // This service may no longer exist?
        console.error("Call to wu-quotes.apple.com produced an error: "+e);
        // Let's just assume the market is open until this can be debugged.
        PropertiesService.getScriptProperties().setProperty("market_open", true);
      }
    }
    if (current_date.getDay() > 0 && current_date.getDay() < 6 && current_date.getHours() == 16 && current_date.getMinutes() == 0) {
      // If the market was open today...
      if (PropertiesService.getScriptProperties().getProperty("market_open") == true) {
        // Update the sheets with today's closing prices.
        PropertiesService.getScriptProperties().setProperty("ticker_index", 0);
      }
    }
  }
}
function TRY_AGAIN(index) {
  // Today's closing price wasn't retrieved successfully. Try again up to 300 times.
  var tries = parseInt(PropertiesService.getScriptProperties().getProperty("ticker_tries"));
  if (!isNaN(tries)) {
    tries++;
  } else {
    tries = 1;
  }
  if (tries == 300) {
    // We tried 300 times. Give up and try again tomorrow.
    PropertiesService.getScriptProperties().deleteProperty("ticker_index");
    PropertiesService.getScriptProperties().deleteProperty("ticker_tries");
  }
  PropertiesService.getScriptProperties().setProperty("ticker_tries", tries);
  PropertiesService.getScriptProperties().setProperty("ticker_index", index); 
}