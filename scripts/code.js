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
