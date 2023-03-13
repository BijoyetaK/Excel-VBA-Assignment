# Excel-VBA-Assignment
Multiple year stock analysis
This workbook on the Click of Calculate button on the Overview sheet, runs a procedure reading through all the Ticker records from 2018 to 2020(each sheet representing records from those years), to return the Opening Stock Price and Closing Stock Price per Ticker for each year, Yearly Change in Stock price, Percentage Change in Stock price and Total Stock volume per Ticker. The same worksheet will also display the Ticker(s) having 'Greatest % Yearly Increase/Decrease and ther Greatest Volume. 
With the click of Reset button on the Overview sheet, it will allow all the sheets to reset the calculated data.(Run Process script) 
This VBA scripting uses the power of Dictionaries as a Vlookup. It allows to use a key value pair thus making it easier to fetch unique Tickers. Key in this case is the Ticker symbol and value is its Open Price, Close Price and Total Volume. For each current Ticker, this procedure adds on to calculate Total Stock volume until the next Ticker. The TickerVolume dictionary hence would contain the Unique Ticknd ResetAllSheets scriptser and their Total Volume alongside the other values. 
Based on the +Change and - Change , the colors of the cells will indicate Green and Red as also Arrows. 
RunAllSheets is used to run the same process across all the sheets and ResetAllSheets script is used to reset the calculate data. 
Sites used for reference: https://www.mrexcel.com/board/threads/is-there-a-way-to-apply-a-macro-across-all-worksheets.997398/ 
For Dictionary: https://www.automateexcel.com/vba/dictionary/, https://stackoverflow.com/questions/36044556/quicker-way-to-get-all-unique-values-of-a-column-in-vba To Enable Dictionary feature in your excel, Go to VBA IDE, Tools -> References - > Check the box for Microsoft Scripting Runtime and Click OK. 
For icons: https://www.bluepecantraining.com/portfolio/excel-vba-apply-icon-set-conditional-formatting-with-vba-macro/
For reset: https://www.wallstreetmojo.com/vba-clear-contents/#h-loop-through-all-the-worksheets-and-clear-contents-of-specific-range
