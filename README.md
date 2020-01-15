# stockpriceupdater
Finally, a reliable way in Google Apps Script to update stock prices.

There are many share price updaters out there, but they tend to be only for US stocks. As someone who has a portfolio crossing USA, Singapore and Hong Kong stocks, the usual techniques don’t work for me because:

* Simply using =GOOGLEFINANCE(ticker) fails for non USA stocks.
* How do you do the correct accounting across SGD and HKD?
* In trying to webscrape prices using a simple URL, sometimes the calls to a site fail, resulting in #N/A (more on this below).

There are three sections to this guide:

1. looking up the prices in a reliable manner.
2. tracking transactions as buy/sell/dividend
3. summarising the transactions

# Price Updater

As mentioned above there are several limitations to the simple approach of `=GOOGLEFINANCE(ticker)` in a cell formula. The first limitation is that [googlefinance doesn’t include SGX](https://investmentmoats.com/money-management/updates-free-google-stock-portfolio-tracker/).

**Table 1: GOOGLEFINANCE works only for USA and HKEX stock tickers, not Singapore nor the Shanghai Hong Kong Stock Connect.**

| Ticker    | Name                       | Formula              | Price   |
|-----------|----------------------------|----------------------|---------|
| VTI       | Vanguard Total USA ETF     | =GOOGLEFINANCE\(A2\) | 166\.59 |
| ES3       | SPDR Singapore ETF         | =GOOGLEFINANCE\(A3\) | 0\.01   |
| ES3\.SI   | SPDR Singapore ETF         | =GOOGLEFINANCE\(A4\) | \#N/A   |
| SGX:ES3   | SPDR Singapore ETF         | =GOOGLEFINANCE\(A5\) | \#N/A   |
| HKEX:2388 | BOC\(hong kong\)           | =GOOGLEFINANCE\(A6\) | \#N/A   |
| 2388      | BOC\(hong kong\)           | =GOOGLEFINANCE\(A7\) | 28\.1   |
| 601318    | Ping An Insurance \(XSSC\) | =GOOGLEFINANCE\(A8\) | \#N/A   |

So one workaround is to use a webscraping technique in Google Sheets:

* `=IMPORTXML("https://sg.finance.yahoo.com/quote/" & TICKER,"//span[@class='Trsdu(0.3s) Trsdu(0.3s) Fw(b) Fz(36px) Mb(-4px) D(b)']" )`

* `=value(mid(IMPORTXML("http://www.dividends.sg/view/" & TICKER, "/html/body/div/div[2]/div/div/div[1]/h4/span"),5,99))`

But [dividends.sg](dividends.sg) didn’t like it and if you access the site too frequently you can get blocked. And scraping [finance.yahoo.com](finance.yahoo.com) sometimes yielded #N/A errors with no explanation on the cell except it failed to get any values. While this is suitable for now-and-then analyses, a single #N/A error will propagate and mean that your whole portfolio calculation will also result in a #N/A error. If you try to use error handling, eg:

`=IFERROR(stock_price, value_if_error)`

Then because you're replacing with some arbitrary value like 0, your portfolio will be inaccurately portrayed. I had to wait for 5 minutes for the errors to refresh and they eventually go away, but the errors didn't always disappear even after waiting.

**So I needed a more robust system.**

## Key Considerations:

* I don’t need it to be updated to seconds because I’m not that kind of trader, but this is just for a daily update on the portfolio. If I really need finer analyses I use the direct numbers on Yahoo Finance or Webull. Therefore, hours to update is fine. We will use the Google App Scripts because it has Triggers that run your script every X hours.
* Must be resistant to website lookup errors, in order to get a timely calculation of portfolio.
* Self-contained on Google Sheets so I don’t have to run anything on a client and I can access it on the go.

## The Solution

Having laid out these specifications, one way to solve them is to use a buffered system for the stock prices. A buffered system means that in the absence of updated prices, you still can fall back to prices that were found earlier. The buffer is a database of prices that were previously successfully looked up.

See Figure 1 for the flowchart of how this buffered system would work.

![Flowchart](/images/F01-flowchart.png)
**Figure 1. Flowchart of a buffered price updating system that updates the stock price *only* if it has successfully found fresh prices.**

## How? The Nuts and Bolts

1. The first thing to do is to create a new sheet in **Google Sheets** and name it “buffer”, and save the file.

2. Then take note of the “url” of your sheet by opening it in your browser and copying the link:
    * If your sheet is at `https://docs.google.com/spreadsheets/d/1jhgjhg9jhgjhgMVjFJ-9WjjhjVEHY/edit`
    * Then your “url” is `1jhgjhg9jhgjhgMVjFJ-9WjjhjVEHY`
    
3. Now navigate to Tools > Script Editor and create a new script and name it “buffer stock prices”. Google App Scripts uses Javascript so we will be coding in that.

4. The code proper. I will show the whole code first and then explain each section.

```
/*
#### Code.js ####

Steps needed:
1. get the unique list of stock symbols and the exchange name
2. attempt to look up the prices
3. if prices are found, update a static database
4. if price is not found, just use the previous value, ie don't update.
*/

// THE MAIN FUNCTION
function updateMasterList() {
  
  // SECTION A - GET SHEET OBJECT
  
  // specify the "url" of your workbook
  var sid = "1Qm9sYdR4qHtsEnMVgJt3ServxtxxS0KFJ-9WYFjVEHY";
  // assign the workbook to an object
  var wb = SpreadsheetApp.openById(sid);
  // get the sheet named "buffer"
  var sheet = wb.getSheetByName("buffer");
  // get the lastrow using a spreadsheet formula on the buffer sheet
  var lastrow = parseInt(sheet.getRange(1, 11).getValue());  
  // set the timestamp to the present time.
  sheet.getRange(2, 11).setValue(getTimeNow());
  
  //SECTION B - GET INPUTS
  
  // these inputs take the format of a 2D array [[s1],[s2],[s3],...]
  var symbols = sheet.getRange(2, 1, lastrow-1, 1).getValues();
  var exchange = sheet.getRange(2, 4, lastrow-1, 1).getValues();    
  
  // Make a record of the input values for debugging purposes.
  Logger.log(symbols);
  Logger.log(exchange);  
  
  //SECTION C - INITIALISE INTERMEDIATE AND OUTPUT VARIABLES
  
  var prices = Array(symbols.length);
  var current_symbol = '';
  var current_exchange = '';
  var url = '';
  var quote = 0.0     
  
  //SECTION D - LOOP THROUGH SYMBOLS AND GET PRICE FROM YAHOO FINANCE
  
  for (var i = 0; i < symbols.length; i++) {
    // Convert each value to string, where integers due to Hong Kong stock numbers gets converted to string wihtout the decimal.
    current_symbol = String(symbols[i]);
    current_exchange = String(exchange[i]);
    
    url = URLFromExchange(current_symbol, current_exchange);
    // if lookupQuote produces an error, quote = -1
    quote = lookupQuote(current_symbol, url);
    prices[i] = [quote];
    
  }
  
  Logger.log([symbols,prices]);
  
  
  // update the sheet in column F, "new_price"  
  Logger.log(prices);
  sheet.getRange(2, 6, lastrow-1, 1).setValues(prices);
  
  // if result is not -1 then update the database
  for (var k = 0; k < prices.length; k++) {
    if (prices[k] != -1) {
    sheet.getRange(2+k, 7).setValue(parseFloat(prices[k]))
    }
  }
  

}


function URLFromExchange(symbol, exchange) {  
  var suffix = ''
  
  switch(exchange) {
    case "USA":
      suffix = '';
      break;
    
    case "SGX":
      suffix = '.SI';
      break;
    
    case "HKEX":
      suffix = '.HK';
      break;
    
    case "XSSC":
      suffix = '.SS';
      break;
  }
    
    
  return 'https://finance.yahoo.com/quote/' + symbol + suffix + '?p=' + Math.floor(Math.random()*1000);
    
}


function lookupQuote(symbol, url) {  
  
  var options = {
    'method':'GET',
    'muteHttpExceptions': true
  };
  
  try {
    for (var n=0;n<3;n++){      
      var page = UrlFetchApp.fetch(url,options);
      if (page.getResponseCode() == 200){
        Logger.log("Attempt " + (n+1))
        break;}
    }
    if (page.getResponseCode() != 200){
      throw "Page failed to load even after 3 attempts."
    }
    
  }
  catch(err) {
    Logger.log("While looking up symbol " + symbol)
    Logger.log("Error occured " + err)
    return -1
  }
  var html = page.getContentText();
  var fromText = '<div class="D(ib) Va(m) Maw(65%) Ov(h)" data-reactid="32">';
  var toText = '</span>';
  
  var initMatch = /(<span class="Trsdu)(.*?)(\<\/span\>)/.exec(html, 1);
  Logger.log(initMatch[0]);
  var finalMatch = /(\>)([\d.,]*)(\<)/.exec(initMatch[0],1)[0];
  var cleanedMatch = parseFloat(finalMatch.replace('<','').replace('>','').replace(',',''));
  
  return cleanedMatch;
}


function testwrite() {
  var sid = "1Qm9sYdR4qHtsEnMVgJt3ServxtxxS0KFJ-9WYFjVEHY";
  var wb = SpreadsheetApp.openById(sid);
  var sheet = wb.getSheetByName("buffer");
  var lastrow = parseInt(sheet.getRange(1, 11).getValue());
  
  sheet.getRange(2,6,5,1).setValues([[1],[2],[3],[4],[5]])

}

function getTimeNow() {  
  var t = new Date();
  var today = t.toLocaleTimeString() + ' ' + t.toLocaleDateString();
  Logger.log(today);
  return today;
  }
```
