// Source: https://youtu.be/_KJsx7QD6dM, https://github.com/kohjb/AmeritradeAPIwGoogleScripts

// ******************************** SPREADSHEET FUNCTIONS ********************************

/**
 * Return the close price of a referenced stock symbol.
 *
 * @param {stockSymbol} string Cell of symbol in caps.
 * @returns The stock's close price.
 *
 * @customfunction
 */
function amtdClosePrice(stockSymbol) {
  var authorization = amtdGetBearerString();
  var options = {
    method: "GET",
    headers: { Authorization: authorization },
    apikey: apikey,
  };
  var myurl =
    "https://api.tdameritrade.com/v1/marketdata/" + stockSymbol + "/quotes";
  var result = UrlFetchApp.fetch(myurl, options);

  //Parse JSON
  var contents = result.getContentText();
  var json = JSON.parse(contents);
  var stock = JSON.stringify(json[stockSymbol]);

  var closePrice = stock["closePrice"];

  return closePrice;
}

/**
 * Returns the positions in your Ameritrade account portfolio.
 *
 * @returns The positions in your portfolio.
 *
 * @customfunction
 */
function amtdPositions() {
  var authorization = amtdGetBearerString();
  var options = {
    method: "GET",
    headers: { Authorization: authorization },
  };
  var extraOptions = "?fields=positions";
  var accountId = "yourAccountID";
  var myUrl =
    "https://api.tdameritrade.com/v1/accounts/" + accountId + extraOptions;
  var result = UrlFetchApp.fetch(myUrl, options);

  //Parse JSON
  var contents = result.getContentText();
  var json = JSON.parse(contents);
  var positions = json["securitiesAccount"]["positions"];

  // I'm interested in these attributes, here you could add your own. (Check the available one's on the API's response example)
  var attributes = [
    "instrument", // The stock symbol is inside this returned "instrumet" object.
    "shortQuantity",
    "longQuantity",
    "averagePrice",
    "marketValue",
  ];

  var array = [];
  var item = [];

  for (var stocki = 0; stocki < positions.length; stocki++) {
    // Iterate over all returned positions
    for (var attributei = 0; attributei < attributes.length; attributei++) {
      // Iterate over the wanted attributes to find the corresponding value in the returned positions
      if (attributes[attributei] === "instrument") {
        // Exemption (maybe a bit nasty) to save the BRK.B symbol as BRK/B to use other services that require it to be in this format
        // You could easily remove this if else statement leaving only the first part.
        if (positions[stocki][attributes[attributei]]["symbol"] !== "BRK.B") {
          item.push(positions[stocki][attributes[attributei]]["symbol"]);
        } else {
          item.push("BRK/B");
        }
      } else {
        item.push(positions[stocki][attributes[attributei]]);
      }
    }
    array.push(item);
    item = [];
  }

  array.sort(function(a, b) {
    // To sort the positions by market value
    return b[4] - a[4];
  });

  return array;
}

/**
 * Returns the cash balance in your Ameritrade account.
 *
 * @returns The cash balance.
 * @customfunction
 */
function amtdCash() {
  var authorization = amtdGetBearerString();
  var options = {
    method: "GET",
    headers: { Authorization: authorization },
  };
  var extraOptions = "?fields=positions";
  var accountId = "yourAccountID";
  var myUrl =
    "https://api.tdameritrade.com/v1/accounts/" + accountId + extraOptions;
  var result = UrlFetchApp.fetch(myUrl, options);

  //Parse JSON
  var contents = result.getContentText();
  var json = JSON.parse(contents);
  var cash = json["securitiesAccount"]["currentBalances"]["cashBalance"];

  return cash;
}

/**
 * Returns the account balance in your Ameritrade account.
 *
 * @returns The account balance.
 * @customfunction
 */
function amtdBalance() {
  var authorization = amtdGetBearerString();
  var options = {
    method: "GET",
    headers: { Authorization: authorization },
  };
  var extraOptions = "?fields=positions";
  var accountId = "yourAccountID";
  var myUrl =
    "https://api.tdameritrade.com/v1/accounts/" + accountId + extraOptions;
  var result = UrlFetchApp.fetch(myUrl, options);

  // Parse JSON
  var contents = result.getContentText();
  var json = JSON.parse(contents);
  var balance = json["securitiesAccount"]["currentBalances"]["equity"];

  return balance;
}

// ******************************** API AUTHENTICATION ********************************

// ********************** LOGIN UI **********************

function amtdShowPane() {
  //Open a SidePane asynchronously. The html will return by calling the function amtdbackfromPane

  linkURL =
    "https://auth.tdameritrade.com/auth?response_type=code&redirect_uri=https%3A%2F%2F127.0.0.1&client_id=" +
    apikey +
    "%40AMER.OAUTHAP";
  var html = HtmlService.createTemplateFromFile("amtd_SidePane").evaluate();
  SpreadsheetApp.getUi().showSidebar(html);
}

function amtdbackfromPane(d) {
  //Called after user clicks Step 2 button on SidePane, return here with dictionary d

  amtdGetTokens(d.returnURI);
}

function amtdGetBearerString() {
  //Call amtd get access token using the rfresh token - check validity of both access and refresh tokens.

  var refresh_token = userProperties.getProperty("refresh_token");
  var refresh_time = userProperties.getProperty("refresh_time");
  var access_token = userProperties.getProperty("access_token");
  var access_time = userProperties.getProperty("access_time");
  var mynow = new Date();

  if (Date.parse(mynow) - Date.parse(access_time) < 29 * 60 * 1000) {
    //Access token is still not expired
    return "Bearer " + access_token;
  } else if (
    Date.parse(mynow) - Date.parse(refresh_time) >
    90 * 24 * 60 * 60 * 1000
  ) {
    //Refresh token expired
    //re-authenticate - amtdshowPane() ?
    return "Re-authentication needed!";
  }

  var formData = {
    grant_type: "refresh_token",
    refresh_token: refresh_token,
    client_id: apikey,
  };
  var options = {
    method: "post",
    payload: formData,
  };
  var myurl = "https://api.tdameritrade.com/v1/oauth2/token";
  var result = UrlFetchApp.fetch(myurl, options);

  //Parse JSON
  var contents = result.getContentText();
  var json = JSON.parse(contents);

  access_token = json["access_token"];
  userProperties.setProperty("access_token", access_token);
  userProperties.setProperty("access_time", access_time);

  if (json.hasOwnProperty("refresh_token")) {
    refresh_token = json["refresh_token"];
    userProperties.setProperty("refresh_token", refresh_token);
    userProperties.setProperty("refresh_time", refresh_time);
  }

  return "Bearer " + access_token;
}

function amtdGetTokens(s) {
  //Receive the URI, strip out the code, and call Ameritrade to receive Bearer Token and RefreshToken
  mycode = decodeURIComponent(s.substring(s.indexOf("code=") + 5));

  var formData = {
    grant_type: "authorization_code",
    access_type: "offline",
    code: mycode,
    client_id: apikey,
    redirect_uri: "https://127.0.0.1",
  };
  var options = {
    method: "post",
    payload: formData,
  };
  var myurl = "https://api.tdameritrade.com/v1/oauth2/token";
  var result = UrlFetchApp.fetch(myurl, options);

  // Logger.log(result);

  //Parse JSON
  var contents = result.getContentText();
  var json = JSON.parse(contents);

  access_token = json["access_token"];
  refresh_token = json["refresh_token"];

  userProperties.setProperty("access_token", access_token);
  userProperties.setProperty("access_time", new Date());
  userProperties.setProperty("refresh_token", refresh_token);
  userProperties.setProperty("refresh_time", new Date());
}

// FOR DEBUGGING
// function amtdputTokens() {
//   //put the access and rerfresh tokens and their times from userProperties in the spreadsheet

//   var tokensheet = "API";
//   var rngAccessToken = "D9";
//   var rngRefreshToken = "D10";

//   var currentssht = SpreadsheetApp.getActive();
//   var sourcesht = currentssht.getSheetByName(tokensheet);

//   var access_token = userProperties.getProperty("access_token");
//   var access_time = userProperties.getProperty("access_time");
//   var refresh_token = userProperties.getProperty("refresh_token");
//   var refresh_time = userProperties.getProperty("refresh_time");

//   sourcesht.getRange(rngAccessToken).setValue(access_token);
//   sourcesht
//     .getRange(rngAccessToken)
//     .offset(0, -1)
//     .setValue(access_time);
//   sourcesht.getRange(rngRefreshToken).setValue(refresh_token);
//   sourcesht
//     .getRange(rngRefreshToken)
//     .offset(0, -1)
//     .setValue(refresh_time);
// }
