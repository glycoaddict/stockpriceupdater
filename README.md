# Detailed tutorial to make your own stock price updater using Google Sheets and Google Apps Scripts

## Summary
Finally, a reliable way using Google Sheets and Google Apps Scripts to update stock prices for a basket comprising stocks from multiple exchanges that are not supported by Sheet's native `=GOOGLEFINANCE` function. Uses web scraping and buffering stock quotes, in order to reliably deliver stock prices.

## Why?
This is a detailed tutorial because it was so painful to build this so I figured it might be useful to some of you out there. 

As I discovered, there are many share price updaters out there, but they tend to be only for US stocks. For someone who wants to lookup a portfolio crossing USA, Singapore and Hong Kong stocks, the usual techniques don’t work because:

* Simply using =GOOGLEFINANCE(ticker) fails for non USA stocks.
* In trying to lookup prices from a given website, sometimes the calls to a site fail, resulting in #N/A (more on this below). Also the native lookup in Sheets makes far too many requests to sites, which problem affects their traffic but if you use Google Sheets it seems to trigger every few minutes and it can't be controlled. 

My intent here is to make a theoretical script useful for looking up stock prices for 10-20 stocks once every 6 hours so this isn't the kind of real time spamming that would be illegal. Not for trading (who would be so stupid as to use hours-old prices!), but just for instructional purposes as this was an interesting problem in web look ups and system design.

This tutorial is step one of three steps. The other two will come at a later time.

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

So one workaround is to use a look up technique in Google Sheets:

* `=IMPORTXML("https://sg.finance.yahoo.com/quote/" & TICKER,"//span[@class='Trsdu(0.3s) Trsdu(0.3s) Fw(b) Fz(36px) Mb(-4px) D(b)']" )`

* `=value(mid(IMPORTXML("http://www.dividends.sg/view/" & TICKER, "/html/body/div/div[2]/div/div/div[1]/h4/span"),5,99))`

But [dividends.sg](dividends.sg) didn’t like it and if you access the site too frequently. And scraping [finance.yahoo.com](finance.yahoo.com) sometimes yielded #N/A errors with no explanation on the cell except it failed to get any values. While this is suitable for now-and-then analyses, a single #N/A error will propagate and mean that your whole portfolio calculation will also result in a #N/A error. If you try to use error handling, eg:

`=IFERROR(stock_price, value_if_error)`

Then because you're replacing with some arbitrary value like 0, your portfolio will be inaccurately portrayed. I had to wait for 5 minutes for the errors to refresh and they eventually go away, but the errors didn't always disappear even after waiting.

Also, it isn't fair to Yahoo if we have a hundred stock prices pinging them every few minutes. But there's no way to control the frequency! Or is there?

**So I needed a more robust system.**

## Key Considerations:

* I don’t need it to be updated to seconds because I’m not that kind of trader, but this is just for a once or twice daily update on a portfolio. If I really need finer analyses I use the direct numbers looking at the Yahoo Finance website manually. Therefore, hours to update is fine. We will use the Google App Scripts because it has Triggers that run your script every X hours.
* Must be resistant to website lookup errors, in order to get reliable calculation of portfolio.
* Self-contained on Google Sheets so I don’t have to run anything on a client and I can access it on the go.

## The Solution

Having laid out these specifications, one way to solve them is to use a buffered system for the stock prices. A buffered system means that in the absence of updated prices, you still can fall back to prices that were found earlier. The buffer is a database of prices that were previously successfully looked up.

See Figure 1 for the flowchart of how this buffered system would work.

![Flowchart](/images/F01-flowchart.png)
**Figure 1. Flowchart of a buffered price updating system that updates the stock price *only* if it has successfully found fresh prices.**

## How? The Nuts and Bolts

1. The first thing to do is to create a new sheet in **Google Sheets** and name it “buffer”, and save the file. Lay out your symbols, exchange, stock name, currency, etc as follows:

| \. | A            | B     | C                 | D        | E        | F          | G               | H | I | J        | K                               |
|----|--------------|-------|-------------------|----------|----------|------------|-----------------|---|---|----------|---------------------------------|
| 1  | stock symbol | Type  | stock name        | exchange | currency | new\_price | buffered\_price |   |   | LASTROW= |                               |
| 2  | 2388         | Stock | BOC\(hong kong\)  | HKEX     | HKD      |            |                 |   |   | LASTRUN= |  |
| 3  | 601318       | Stock | Ping An Insurance | XSSC     | CNH      |            |                 |   |   |
| 4  | AW9U         | Reit  | First Reit        | SGX      | SGD      |            |                 |   |   |
| 5  | BLCM         | Stock | Bellicum          | USA      | USD      |            |                 |   |   |
| 6  | CRPU         | Reit  | Sasseur REIT      | SGX      | SGD      |            |                 |   |   |
| 7  | DIS          | Stock | Disney            | USA      | USD      |


2. Now navigate to Tools > Script Editor and create a new script and name it “buffer stock prices”. Google App Scripts uses Javascript so we will be coding in that.

3. The code proper. I will show the whole code first and then explain each section.

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
  
  // SUBSECTION D1
    
  for (var i = 0; i < symbols.length; i++) {
  // Convert each value to string, where integers due to Hong Kong stock numbers gets converted to string without the decimal.
    current_symbol = String(symbols[i]);
    current_exchange = String(exchange[i]);
    
    // Get URL based on which stock exchange
    url = URLFromExchange(current_symbol, current_exchange);
    
    // Look up the quote in Yahoo Finance
    // if lookupQuote produces an error, quote will return -1
    quote = lookupQuote(current_symbol, url);
    
    // set quote value in the (output) prices array
    prices[i] = [quote];
    
  }
  
  Logger.log([symbols,prices]);  
  
  // SUBSECTION D2 - UPDATING THE VALUES
  
  // update the sheet in column F, under the header "new_price"
  sheet.getRange(2, 6, lastrow-1, 1).setValues(prices);
  
  // if result is not -1, ie a correct quote was found, then update the database
  for (var k = 0; k < prices.length; k++) {
    if (prices[k] != -1) {
    sheet.getRange(2+k, 7).setValue(parseFloat(prices[k]))
    }
  }
  

}


// FUNCTION: ALTER THE URL'S SUFFIX BASED ON WHICH EXCHANGE IS BEING INVOKED
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

// LOOK UP THE QUOTE WITH 3 ATTEMPTS
function lookupQuote(symbol, url) {  
    
  var options = {
    'method':'GET',
    // muteHttpExceptions will prevent throwing an error if html errors such as 404 are encountered.
    'muteHttpExceptions': true
  };
  
  try {
    // Three attempts to load the page
    for (var n=0;n<3;n++){      
      var page = UrlFetchApp.fetch(url,options);
      // code 200 means successfully loaded.
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
  
  // get the html as string
  var html = page.getContentText();
  
  // use series of regex searches to extract out the desired strings 
  // regex 1
  var initMatch = /(<span class="Trsdu)(.*?)(\<\/span\>)/.exec(html, 1);
  Logger.log(initMatch[0]);  
  // regex 2
  var finalMatch = /(\>)([\d\.\,]*)(\<)/.exec(initMatch[0],1)[0];
  // remove any commas or <> to arrive at just the numbers and coerce into Float.
  var cleanedMatch = parseFloat(finalMatch.replace('<','').replace('>','').replace(',',''));
  
  return cleanedMatch;
}

// GET THE DATETIMESTAMP
function getTimeNow() {  
  var t = new Date();
  var today = t.toLocaleTimeString() + ' ' + t.toLocaleDateString();
  Logger.log(today);
  return today;
  }
```

**Let's break it down**

## SECTION A - GET SHEET OBJECT

```
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
```

The next thing is to take note of the **identifier** of your sheet by opening it in your browser and copying the link:
    * If your sheet is at `https://docs.google.com/spreadsheets/d/1jhgjhg9jhgjhgMVjFJ-9WjjhjVEHY/edit`
    * Then your “url” is `1jhgjhg9jhgjhgMVjFJ-9WjjhjVEHY`
    * we will use this to write to your Google Sheet.
    * IMPORTANT: Google will prompt you to give your Apps Script account permission to change your Drive contents. Check the permissions and accept it. *Basically you need permission to read and write your own Sheet because Google sees your .js script as an external "addin" and therefore has security measures.*

Update the `var sid =` with your sheet's identifider (enclosed in quotes because it is a string).

The only thing unclear in SECTION A is:

```
  // get the lastrow using a spreadsheet formula on the buffer sheet
  var lastrow = parseInt(sheet.getRange(1, 11).getValue());  
```

For our purposes we need to be able to specify the last row on the sheet. We will use a spreadsheet-based formula, so set your K1:L2 cells as follows (**Table 2**). The LASTRUN value in L2 will be used by the script so we can keep track of when it was last updated.

**Table 2: Lastrow formula**
| row\\column | J        | K                                             |
|-------------|----------|-----------------------------------------------|
| 1           | LASTROW= | =rows\(filter\(A:A, not\(ISBLANK\(A:A\)\)\)\) |
| 2           | LASTRUN= |

Explaining the formula `=rows(filter(A:A, not(ISBLANK(A:A))))`:

1. `filter(A:A, not(ISBLANK(A:A)))` apply a filter to column A to find non blank cells. i.e., get an array of cells with values. **Note that this assumes no blanks in between the rows, which we just have to keep in mind but this is a simple constraint to fulfil.**

2. `=rows()` counts the number of rows in the filtered result, and assuming no blanks, this is the row number of the last value in the sheet.

Check that your `LASTROW` value is now properly calculated. You can experiment with blanks too to see the effect.

Now we can simply lookup the "buffer" sheet in your Google Sheets object, and get the `LASTROW` value in **K1 (row=1, col=11)** thus:

```
  // get the lastrow using a spreadsheet formula on the buffer sheet
  var lastrow = parseInt(sheet.getRange(1, 11).getValue());  
```

   where `sheet` is the "buffer" sheet object; `.getRange(row, column)` selects the cell of interest; and `.getValue()` extracts the value of that cell. The data type appears to be automatically picked, which in this case is a float.

Next, to get the timestamp, we have to first define a separate function, as follows.

### getTimeNow() function

You need to create a timestamp so you know when the numbers came from, to prevent you using old data. Javascript feels more difficult to generate timestamps because of its enormous flexibility. One way to get a string of the date and time is the following function placed in your script. I used the combination of the two functions  `.toLocaleTimeString()` and `.toLocaleDateString()` because no matter how I tried to play around with the options, I couldn’t get a single Date object to show *both* the DATE and TIME.

```
// GET THE DATETIMESTAMP
function getTimeNow() {  
  var t = new Date();
  var today = t.toLocaleTimeString() + ' ' + t.toLocaleDateString();
  Logger.log(today);
  return today;
  }
```

With this added to the bottom of your script and **NOT INSIDE YOUR MAIN FUNCTION**, the calling line simply gets a string of the date and time stamp and writes it to **K2 (row=2, column=1)**:

```
  // set the timestamp to the present time.
  sheet.getRange(2, 11).setValue(getTimeNow());
```

## SECTION B - GET INPUTS

```
//SECTION B - GET INPUTS
  
  // these inputs take the format of a 2D array [[s1],[s2],[s3],...]
  var symbols = sheet.getRange(2, 1, lastrow-1, 1).getValues();
  var exchange = sheet.getRange(2, 4, lastrow-1, 1).getValues();    
  
  // Make a record of the input values for debugging purposes.
  Logger.log(symbols);
  Logger.log(exchange);  
```

We use `.getRange(2, 1, lastrow-1, 1)` based on `.getRange(row, column, row_size, column_size)` where the parameter `row_size` means how many rows to collect, which in this case is one less than the number of rows to account for the header (note that we started from row 2); and `column_size` of `1` to specify only one column. Because we are selecting a range, we will need to extract the values slighly differently.

In the same way as getting the values from a particular range above, we use a variant `.getValues()`, which returns the range values as a 2D array. The data takes the form of a 2D array even though we have only extracted a single columnnar range. The lenth of the second dimension in this case is 1. Why this is important is that later when we are creating the output array of prices, we need to store them in an array: `[[price1],[price2],...[priceN]]`.

Our inputs are now two arrays of **symbols** and their respective **exchanges**, e.g.:
* `[[2388.0], [601318.0], [AW9U], [BLCM], [CRPU], [DIS], [EDUC], [ES3], [EWM], [FB], [G3B], [GOOG], [J91U], [NXST], [O9P], [OV8], [TEAM], [TMO], [VET], [VOO], [VTI], [XAR]]`
* `[[HKEX], [XSSC], [SGX], [USA], [SGX], [USA], [USA], [SGX], [USA], [USA], [SGX], [USA], [SGX], [USA], [SGX], [SGX], [USA], [USA], [USA], [USA], [USA], [USA]]`


## SECTION C - INITIALISE INTERMEDIATE AND OUTPUT VARIABLES

```
  //SECTION C - INITIALISE INTERMEDIATE AND OUTPUT VARIABLES
  
  var prices = Array(symbols.length);
  var current_symbol = '';
  var current_exchange = '';
  var url = '';
  var quote = 0.0    
```

We have two options when it comes to creating array outputs:

* create a zero-length array and append to the array each time in the loop. This has the danger of insufficient values being added if an error is encountered, which would make the output the wrong size for the Range you want to write the values to later.

* pre-make an array of the desired length (filled with blanks or zeros) and update each position in the array as needed. This ensures the correct size of the array before you start to fill it. This is also likely faster as the append operation is usually slower than an assign operation. The code is `var prices = Array(symbols.length);`

We also initialise four other necessary variables here. Why here? Generally for readability, so we know the gamut of variables before we dive into the loop, but so I don't generate a new variable for each iteration of the loop, which is probably heavier on memory. 

## SECTION D - LOOP THROUGH SYMBOLS AND GET PRICE FROM YAHOO FINANCE

We now look at the heart of this script, the loop that will go through the list of symbols and exchanges and retrieve the stock price.

```
  // SUBSECTION D1
    
  for (var i = 0; i < symbols.length; i++) {
  // Convert each value to string, where integers due to Hong Kong stock numbers gets converted to string without the decimal.
    current_symbol = String(symbols[i]);
    current_exchange = String(exchange[i]);
    
    // Get URL based on which stock exchange
    url = URLFromExchange(current_symbol, current_exchange);
    
    // Look up the quote in Yahoo Finance
    // if lookupQuote produces an error, quote will return -1
    quote = lookupQuote(current_symbol, url);
    
    // set quote value in the (output) prices array
    prices[i] = [quote];
    
  }
  
  Logger.log([symbols,prices]);  
  
  // SUBSECTION D2 - UPDATING THE VALUES
  
  // update the sheet in column F, under the header "new_price"
  sheet.getRange(2, 6, lastrow-1, 1).setValues(prices);
  
  // if result is not -1, ie a correct quote was found, then update the database
  for (var k = 0; k < prices.length; k++) {
    if (prices[k] != -1) {
    sheet.getRange(2+k, 7).setValue(parseFloat(prices[k]))
    }
  }
```

## SUBSECTION D1

`current_symbol = String(symbols[i]);`
`current_exchange = String(exchange[i]);`

The first SUBSECTION D1 converts the stock symbols to strings. This is needed because the Hong Kong exchange stock symbols are numbers and are parsed as floats when extracting from the Google Sheet. For example, the stock of Bank of China (Hong Kong)'s symbol is "2388" but is parsed as `2388.0`. Converting back to string renders it properly `"2388"`.

```
    // Get URL based on which stock exchange
    url = URLFromExchange(current_symbol, current_exchange);
```
The above line calls a function that does what it says, namely return the correct form of the yahoo finance url according to which exchange the symbol is listed on.

The function:

```
// FUNCTION: ALTER THE URL'S SUFFIX BASED ON WHICH EXCHANGE IS BEING INVOKED
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
```

Right now this function covers the exchanges:
* SGX
* Hong Kong
* Shanghai Hong Kong Stock Connect
* USA, i.e. NASDAQ + NYSE

How this function works is basically depending on the exchange, it appends a different suffix to the url `'https://finance.yahoo.com/quote/' + symbol + suffix + '?p=' + Math.floor(Math.random()*1000);`. E.g. for Singapore Stock Exchange, SGX it suffixes `".SI"`. And it adds in a spurious query of an integer random number from 0-1000 `?p=Math.floor(Math.random()*1000)` to mix things up in case multiple calls to the same url gets incorrecly flagged by the yahoo servers as some sort of spamming or attack (no idea if they would). So for the symbol `ES3`, the final url will be `https://finance.yahoo.com/quote/ES3.SI?=123`.

The next step is to extract the stock price from the HTML of the loaded website, using this function call below:
```
    // Look up the quote in Yahoo Finance
    // if lookupQuote produces an error, quote will return -1
    quote = lookupQuote(current_symbol, url);
```

The function lookupQuote() is explained below:

```
// LOOK UP THE QUOTE WITH 3 ATTEMPTS
function lookupQuote(symbol, url) {  
    
  var options = {
    'method':'GET',
    // muteHttpExceptions will prevent throwing an error if html errors such as 404 are encountered.
    'muteHttpExceptions': true
  };
  
  try {
    // Three attempts to load the page
    for (var n=0;n<3;n++){      
      var page = UrlFetchApp.fetch(url,options);
      // code 200 means successfully loaded.
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
  
  // get the html as string
  var html = page.getContentText();
  
  // use series of regex searches to extract out the desired strings 
  // regex 1
  var initMatch = /(<span class="Trsdu)(.*?)(\<\/span\>)/.exec(html, 1);
  Logger.log(initMatch[0]);  
  // regex 2
  var finalMatch = /(\>)([\d.,]*)(\<)/.exec(initMatch[0],1)[0];
  // remove any commas or <> to arrive at just the numbers and coerce into Float.
  var cleanedMatch = parseFloat(finalMatch.replace('<','').replace('>','').replace(',',''));
  
  return cleanedMatch;
}
```
This function is a bit longer and more complicated, so I will take it one chunk at a time.

1. The function takes the `symbol`, and also the `url` generated by `URLFromExchange()`:
```
function lookupQuote(symbol, url) {  
```

2. Next we make three tries to load the webpage at `url`. 
    * `var page = UrlFetchApp.fetch(url,options);` attempts to load the page at `url`.
    * Only proceeds if page is loaded successfully, which is signalled by a `page.getResponseCode() == 200`. 
    * If loading is unsuccessful even after 3 tries, this function returns a result of `-1` for the stock price. 
    * `-1` is an impossible stock price and is used to signal later on **NOT** to update *that* particular symbol's price.
    * the statement `try {do something} catch(err) {do another thing}` means to "try" to do something but if an error occurs, then do another thing. 
  
```
  try {
    // Three attempts to load the page
    for (var n=0;n<3;n++){      
      var page = UrlFetchApp.fetch(url,options);
      // code 200 means successfully loaded.
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
```

3. Next we extracts the html of the loaded page as a string using `var html = page.getContentText();`.

**When web scraping we need to know what section exactly do we want to extract information from. This is what the page looks like:**

![Screenshot of yahoo finance](/images/F02-yahoofinanceexample.png). 

**And if you right click on the stock price of 28.200 and click inspect (in Chrome) to get to the source code, you will see this:**

![inspect source code](/images/F03-inspectedyahoo.png)

One way to extract out the stock price is to use regex to cut out the desired string based on unique patterns before and after the desired string. (Another way would be to parse the website as xml structured data and access the correct tag, e.g. the exact span with a unique class identifier, but I couldn't work out the correct methods to parse the page as xml when loading with `UrlFetchApp.fetch(url,options)`).

4. Extract out a section of interest using Regular Expressions (Regex). Specifically looks in the html for `<span>` with class property starting with `Trsdu`, then matches everything in a non-greedy fasion using `(.*?)` (where the non-greedy operator`?` tells it to take the shortest possible match. All the way until it encounters a `</span>`., which is what closes off the <span> class.

If the code patterns are very complex and you're finding it difficult to identify the unique pattern, it is sometimes easier to make multiple slices, from broad to specific. That way, any unwanted but similar patterns are sliced away in earlier steps, making it easier to zoom in on the exact value of interest.

The two sequential regex steps I used were:

**Regex 1**, where my broadest cut was to take the pattern of `<span class="Trsdu` all the way to the closing `</span>`. Note that in Javascript, regex expressions are enclosed between the two slashes `/regex_goes_here/.exec(string_to_execute_on, number_to_find)` 
```
  // use series of regex searches to extract out the desired strings 
  // regex 1
  var initMatch = /(<span class="Trsdu)(.*?)(\<\/span\>)/.exec(html, 1);  
```
This resulted in `<span class="Trsdu(0.3s) Trsdu(0.3s) Fw(b) Fz(36px) Mb(-4px) D(b)" data-reactid="14">28.200</span>`. 

As you can see, the stock price is `28.200`, as expected from our inspection of the website and the source code. So we are in the right direction. But we still haven't isolated the stock price so a second regex is needed.

**Regex 2**, where I isolate the pattern of `>28.200<` using the below regex. Note that special characters like `<` and `>` require the escape character "\", as in `\<` will match the string `<`. Escape character \ tells the regex that you want the next character as written.
```
  // regex 2
  var finalMatch = /(\>)([\d\.\,]*)(\<)/.exec(initMatch[0],1)[0];
```
Above, the `\d` means digit and the `[\d\.\,]*` means I wish to look for any combination of digits, periods and commas (note the use of the escape character \). Commas handle the thousands separator if in use as in `1,001`. The final `*` tells the regex that I want any number of such characters. 

The result is `>28.200<`. As you can see we still need to remove the <>, and also any commas as the thousands separator because Javascript cannot parse `1,001` as an integer or float i.e. 1001 but will throw an error. To remove these pesky characters, I simply do sequential string replacements using `finalMatch.replace('<','').replace('>','').replace(',','')`, and then convert from string to float using `parseFloat(string)`.

```
  // remove any commas or <> to arrive at just the numbers and coerce into Float.
  var cleanedMatch = parseFloat(finalMatch.replace('<','').replace('>','').replace(',',''));
```  
  
5. finally, return the value of the stock price (now already a float) using `return cleanedMatch;`

That's the end of the custom function `lookupQuote(symbol, url)`.
  
Now let's head back up to SECTION D1, which I reproduce here because you may have forgotten your place:

```
  //SECTION D - LOOP THROUGH SYMBOLS AND GET PRICE FROM YAHOO FINANCE
  
  // SUBSECTION D1
    
  for (var i = 0; i < symbols.length; i++) {
  // Convert each value to string, where integers due to Hong Kong stock numbers gets converted to string without the decimal.
    current_symbol = String(symbols[i]);
    current_exchange = String(exchange[i]);
    
    // Get URL based on which stock exchange
    url = URLFromExchange(current_symbol, current_exchange);
    
    // Look up the quote in Yahoo Finance
    // if lookupQuote produces an error, quote will return -1
    quote = lookupQuote(current_symbol, url);
    
    // set quote value in the (output) prices array
    prices[i] = [quote];
    
  }
```

We had just examined `quote = lookupQuote(current_symbol, url);`

The final step to SUBSECTION D1 is to update the output array `prices` with the value of `quote`. The for loop ensures that we go through each symbol-exchange pair in the input arrays.

## SUBSECTION D2 - UPDATING THE VALUES

```
  // SUBSECTION D2 - UPDATING THE VALUES
  
  // update the sheet in column F, under the header "new_price"
  sheet.getRange(2, 6, lastrow-1, 1).setValues(prices);
  
  // if result is not -1, ie a correct quote was found, then update the database
  for (var k = 0; k < prices.length; k++) {
    if (prices[k] != -1) {
    sheet.getRange(2+k, 7).setValue(parseFloat(prices[k]))
    }
  }
```

We come to the ultimate code block. Just updating the spreadsheet with the quotes you have now obtained. Remember that a quote value of `-1` means we failed to load the webpage? So we will skip any quotes with value of `-1`.

Recall that our "buffer" sheet had the following contents:

| \. | A            | B     | C                 | D        | E        | F          | G               | H | I | J        | K                               |
|----|--------------|-------|-------------------|----------|----------|------------|-----------------|---|---|----------|---------------------------------|
| 1  | stock symbol | Type  | stock name        | exchange | currency | new\_price | buffered\_price |   |   | LASTROW= | 8                              |
| 2  | 2388         | Stock | BOC\(hong kong\)  | HKEX     | HKD      |            |                 |   |   | LASTRUN= |  |
| 3  | 601318       | Stock | Ping An Insurance | XSSC     | CNH      |            |                 |   |   |
| 4  | AW9U         | Reit  | First Reit        | SGX      | SGD      |            |                 |   |   |
| 5  | BLCM         | Stock | Bellicum          | USA      | USD      |            |                 |   |   |
| 6  | CRPU         | Reit  | Sasseur REIT      | SGX      | SGD      |            |                 |   |   |
| 7  | DIS          | Stock | Disney            | USA      | USD      |

`sheet.getRange(2, 6, lastrow-1, 1).setValues(prices);` will place the quotes you found into the range F2:F7.

The next for loop will update the buffered_price column if the quote is not `-1`.

Now execute your script by selecting from the drop down list at the top of your Script editor `updateMasterList` and click the play button.

After it completes running, you can check the console logs by pressing CTRL+ENTER to see what your script did. And then checking the Google Sheets you'll find it has updated thus!

    
| \. | A            | B     | C                 | D        | E        | F          | G               | H | I | J        | K                               |
|----|--------------|-------|-------------------|----------|----------|------------|-----------------|---|---|----------|---------------------------------|
| 1  | stock symbol | Type  | stock name        | exchange | currency | new\_price | buffered\_price |   |   | LASTROW= | 8                              |
| 2  | 2388         | Stock | BOC\(hong kong\)  | HKEX     | HKD      | 28\.2      | 28\.2           |   |   | LASTRUN= | 9:47:59 PM HKT January 15, 2020 |
| 3  | 601318       | Stock | Ping An Insurance | XSSC     | CNH      | 85\.81     | 85\.81          |
| 4  | AW9U         | Reit  | First Reit        | SGX      | SGD      | 1          | 1               |
| 5  | BLCM         | Stock | Bellicum          | USA      | USD      | 1\.75      | 1\.75           |
| 6  | CRPU         | Reit  | Sasseur REIT      | SGX      | SGD      | 0\.91      | 0\.91           |
| 7  | DIS          | Stock | Disney            | USA      | USD      | 145\.2     | 145\.2          |

**And you're done with the coding!**

# Creating a Trigger

Once you're satisfied that your script works properly, from your Scripts editor, go to the menu `Edit > Project Triggers`. Create a new trigger, with these settings:

![Trigger settings](/images/F04-triggersettings.png)

Click save and now your trigger and it will execute to the time period you specified.

You can now use a `=VLOOKUP` function in Google Sheets to use these buffered prices. You'll never encounter a #N/A error again when looking up prices!

# Happy Investing!
