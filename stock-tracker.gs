var tickers = ["SCHB", "SCHA", "SCHX", "SCHF", "SCHE", "SCHD", "MUB", "C", "BRK.B", "F", "AAPL", "APTV", "DLPH"];

function FIND_TODAYS_CELL(sheet_name) {
  var values = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange("A:A").getValues();
  var current_date = new Date();
  current_date.setHours(0,0,0,0)
  var column_count = 0;
  while (values[column_count][0] != "") {
    var cell_date = new Date(values[column_count][0]);
    if (current_date.getTime() == cell_date.getTime()) {
      return (column_count + 1);
    }
    column_count++;
  }
  return (column_count + 1);
}
function GET_REALTIME_PRICING() {
  var current_date = new Date();
  var api_try_again = parseInt(PropertiesService.getScriptProperties().getProperty("api_try_again"));
  // To save on API calls, only run this routine during market hours
  if (current_date.getDay() > 0 && current_date.getDay() < 6 && current_date.getHours() > 9 && (current_date.getHours() < 16 || !isNaN(api_try_again))) {
    // Check to see whether the market is open today
    if (PropertiesService.getScriptProperties().getProperty("market_open") == null) {
      // The IPX API offers this: 
      // https://iexcloud.io/docs/api/#u-s-holidays-and-trading-dates
      // May want to consider switching to this if Apple takes away this functionality
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
    
    // Only proceed with API calls if the market is open
    var market_status = PropertiesService.getScriptProperties().getProperty("market_open");
    if (market_status == "true") {   
      var ticker_str = tickers.join(",");
      var options = {
        headers: {
          "Cache-Control": "max-age=0"
        }
      };
      var response = UrlFetchApp.fetch("https://cloud.iexapis.com/stable/stock/market/batch?symbols="+ticker_str+"&types=quote&token="+PropertiesService.getScriptProperties().getProperty("api_key"), options);
      var json = JSON.parse(response);
      for (var ticker in json) {
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ticker).getRange("F2").setValue(json[ticker]["quote"]["latestPrice"]);
        // Determine if the market has closed
        var closing_date = new Date(json[ticker]["quote"]["closeTime"]);
        if (current_date.getHours() > 15) {
          if (closing_date.getDate() == current_date.getDate()) {
            // The market has closed. Update the individual ticker sheets with closing/historical data for today
            var cell_number = FIND_TODAYS_CELL(ticker);
            if (cell_number !== false) {
              // Historical data is no longer re-built after hours. IEX offers this, but it is a (deprecated?) V1 API call. Look here if data needs to be recovered:
              // https://iexcloud.io/docs/api/#historical-prices
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ticker).getRange(cell_number, 1).setValue(current_date.toLocaleDateString("en-US"));
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ticker).getRange(cell_number, 2).setValue(json[ticker]["quote"]["close"]);
            }
          } else {
            // Need to keep making API calls until all closing data has been received
            TRY_AGAIN();
            // Return false before properties can be reset below
            return false;
          }
        } else {
          // Return false before properties can be reset below
          return false;          
        }
      }
      // Check and update YTD performance (if necessary)
      CALCULATE_YTD_PERFORMANCE();
      // Reset counters to prepare for the next market day
      PropertiesService.getScriptProperties().deleteProperty("market_open");
      PropertiesService.getScriptProperties().deleteProperty("api_try_again");
    }
  }
}
function CALCULATE_YTD_PERFORMANCE() {
  var closing_ytd = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Overview").getRange("K2").getValue();
  var high_ytd = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Overview").getRange("K3").getValue();
  
  if (closing_ytd > high_ytd) {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Overview").getRange("K3").setValue(closing_ytd);
  }
}
function TRY_AGAIN() {
  // Today's closing prices weren't retrieved successfully. Try again up to 300 times.
  var tries = parseInt(PropertiesService.getScriptProperties().getProperty("api_try_again"));
  if (!isNaN(tries)) {
    tries++;
  } else {
    tries = 1;
  }
  if (tries == 300) {
    // We tried 300 times. Give up and try again tomorrow.
    PropertiesService.getScriptProperties().deleteProperty("api_try_again");
  }
  PropertiesService.getScriptProperties().setProperty("api_try_again", tries);
}