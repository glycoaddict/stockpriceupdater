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

![Flowchart](/images/F01-flowchart.png)

**Table 1: GOOGLEFINANCE works only for USA and HKEX stock tickers, not Singapore nor the Shanghai Hong Kong Stock Connect.**

