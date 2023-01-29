var tickers = ["SCHB", "SCHA", "SCHX", "SCHF", "SCHE", "SCHD", "SCHG", "MUB", "C", "BRK.B", "F", "AAPL", "APTV"];

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

  if (current_date.getDay() < 6) {
    // Check to see whether the market is open today. This needs to be done before market open.
    if (PropertiesService.getScriptProperties().getProperty("market_open") == null &&
        (current_date.getDay() > 0 && current_date.getDay() < 6) &&
        (current_date.getHours() == 9 && current_date.getMinutes() < 30)) {
          var options = {
            headers: {
              "Cache-Control": "max-age=0"
            }
          };

          // Retrieve the next market holiday and check to see whether it matches today's date.
          var response = UrlFetchApp.fetch("https://cloud.iexapis.com/stable/ref-data/us/dates/holiday/next/1/?token="+PropertiesService.getScriptProperties().getProperty("api_key"), options);
          var json = JSON.parse(response);
          var today = Utilities.formatDate(new Date(), "GMT-4", "yyyy-MM-dd");

          if (json[0]["date"] == today) {
            // Market is closed
            PropertiesService.getScriptProperties().setProperty("market_open", false);
          } else {
            // Market is open
            PropertiesService.getScriptProperties().setProperty("market_open", true);
          }
    }
    // To save on API calls, only run this routine during market hours. Allow 5 minutes after close to begin collecting closing prices.
    if ((current_date.getDay() > 0 && current_date.getDay() < 6) &&
      ((current_date.getHours() == 9 && current_date.getMinutes() > 30) || current_date.getHours() > 9) &&
      (((current_date.getHours() == 16 && current_date.getMinutes() <= 5) || current_date.getHours() < 16) ||
        !isNaN(api_try_again))) {
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
            var tickers_closed = 0;
            for (var ticker in json) {
              var ticker_price = GET_CURRENT_PRICE(json, ticker);
              if (current_date.getHours() < 16) {
                SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ticker).getRange("F2").setValue(ticker_price);
              }
              // Determine if the market has closed
              if (json[ticker]["quote"]["isUSMarketOpen"] == false) {
                var closing_date = new Date(json[ticker]["quote"]["latestTime"]);
                // If the number of API calls exceed 400, it's likely that one or more tickers never reported its closing price to IEX.
                // Give up here and get the "latest price" for all tickers so that the day can be properly closed out.
                if (parseInt(PropertiesService.getScriptProperties().getProperty("api_try_again")) >= 400) {
                  tickers_closed++;
                } else {
                  if (closing_date.getMonth() == current_date.getMonth() &&
                        closing_date.getDate() == current_date.getDate() &&
                        closing_date.getFullYear() == current_date.getFullYear() &&
                        json[ticker]["quote"]["latestSource"] == "Close") {
                          tickers_closed++;
                  } else {
                    // Need to keep making API calls until all closing data has been received
                    TRY_AGAIN();
                    // Return false before properties can be reset below
                    return false;
                  }
                }
              }
            }
            if (current_date.getHours() > 15) {
              if (tickers_closed == tickers.length) {
                for (var ticker in json) {
                  var ticker_price = GET_CURRENT_PRICE(json, ticker);
                  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ticker).getRange("F2").setValue(ticker_price);
                  // The market has closed. Update the individual ticker sheets with closing/historical data for today
                  var cell_number = FIND_TODAYS_CELL(ticker);
                  if (cell_number !== false) {
                    // Historical data is no longer re-built after hours. IEX offers this, but it is a (deprecated?) V1 API call. Look here if data needs to be recovered:
                    // https://iexcloud.io/docs/api/#historical-prices
                    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ticker).getRange(cell_number, 1).setValue(current_date.toLocaleDateString("en-US"));
                    var ticker_price = GET_CURRENT_PRICE(json, ticker);
                    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ticker).getRange(cell_number, 2).setValue(ticker_price);
                  }
                }
              } else {
                // Need to keep making API calls until all closing data has been received
                TRY_AGAIN();
                // Return false before properties can be reset below
                return false;
              }

              // Check and update YTD performance (if necessary)
              CALCULATE_YTD_PERFORMANCE();
              // Reset counters to prepare for the next market day
              PropertiesService.getScriptProperties().deleteProperty("market_open");
              PropertiesService.getScriptProperties().deleteProperty("api_try_again");
            }
          } else {
            if (current_date.getHours() > 15) {
              // Reset counters to prepare for the next market day
              PropertiesService.getScriptProperties().deleteProperty("market_open");
              PropertiesService.getScriptProperties().deleteProperty("api_try_again");
            }
          }
    }
  }
}
function GET_CURRENT_PRICE(json_obj, ticker) {
  // Sometimes the last IEX price is null. If so, calculate the price by using: ((bid + ask) / 2)
  if (json_obj[ticker]["quote"]["latestPrice"] == null) {
    if (json_obj[ticker]["quote"]["iexBidPrice"] != null && json_obj[ticker]["quote"]["iexAskPrice"] != null &&
        json_obj[ticker]["quote"]["iexBidPrice"] > 0 && json_obj[ticker]["quote"]["iexAskPrice"] > 0) {
      return ((json_obj[ticker]["quote"]["iexBidPrice"] + json_obj[ticker]["quote"]["iexAskPrice"]) / 2);
    } else {
      // If the last IEX price and the bid/ask prices are null, use the last closing price
      return json_obj[ticker]["quote"]["previousClose"];
    }
  } else {
    return json_obj[ticker]["quote"]["latestPrice"];
  }
}
function CALCULATE_YTD_PERFORMANCE() {
  var closing_ytd = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Overview").getRange("K2").getValue();
  var high_ytd = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Overview").getRange("K3").getValue();

  if (closing_ytd > high_ytd) {
    var current_date = new Date();
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Overview").getRange("K3").setValue(closing_ytd);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Overview").getRange("L3").setValue(current_date.toLocaleDateString("en-US"));
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
  if (tries == 1000) {
    // We tried 1000 times. Give up and try again tomorrow.
    PropertiesService.getScriptProperties().deleteProperty("api_try_again");
    return false;
  }
  PropertiesService.getScriptProperties().setProperty("api_try_again", tries);
}
