# getRoic
Collection of scripts useful to know the health status of a company and its fair value. It is based on the yfinance library.
- Dividend Discounted Model 
- Discounted Cash Flow
- Stock Picking Model, based on balance sheet multiples

## Dependencies
- selenium
- pandas
- numpy
- yfinance
- openpyxl
- pyfiglet

# How it works

1. [Industry Table](#finviz)
2. [Selenium Settings](#selenium)
3. [Warnings](#warnings)

## Industry Table <a name="finviz"></a>
All scripts take the names of the companies you have selected from an excel sheet, python processes the data and then saves the results in a new sheet.

balance-sheet-multiples.py script uses a csv exported from [finviz](https://finviz.com/grp_export.ashx?g=industry&v=120&o=name). 

Next you need to use text-to-columns function to split the text into columns. 

At the end save the file in xlsx format and then you need to indicate the file path in the script.

## Selenium Settings <a name="selenium"></a>

For selenium settings download [chromedriver](https://chromedriver.chromium.org/). 

I used two extensions. The first one to [block ads](https://chrome.google.com/webstore/detail/ublock-origin/cjpalhdlnbpafiamejdnhcphjbkeiagm) and the second one to [bypass cookie acceptance](https://chrome.google.com/webstore/detail/i-dont-care-about-cookies/fihnjjcciajhdojfnbdddfaoknhalnja). 

Take the crx files of these extensions and change the folder path in the script to the folder path where you stored them

## Warnings <a name="warnings"></a>
The project is still in a testing state, known issues are:

- corrupted excel file after forced script closure *(this does not always happen)*
- parameters taken from yfinance not available with some companies
- please note that yahoo may prevent the page from loading due to too many queries. for this reason a timeout of 60s has been added. use a vpn if you plan to make many queries
