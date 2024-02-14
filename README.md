# StockMarketAnalysis
Challenge 2
* Create a script that loops through all the stocks for the given year and outputs the following information:
* The ticker symbol: Methodology used is to go through the rows and remove duplicates and the inspiration came from the instructor providing examples in class.
* The Yearly Change: Methodology used is to go through the rows and whenever there is ticker changes, it will grab the difference between the last ticker's closing price and the first ticker's opening price for the same tickers. I came up with the logic and code myself.
* The percentage change was Yearly Change / opening price of that ticker * 100%. I got the percent sign syntax from ChatGPT but came up with the logic and the rest of the code myself.
* The total stock volume is accumulating the stock volume for the same ticker. It will stop collecting once there is a ticker name change. I came up with the logic and code myself.
* Adding functionality: grabbing the maximum / minimum from the specific data ranges and output the value. Using the same row to locate the corresponding ticker symbol. I got the maximum / minimum syntax from ChatGPT but came up with the logic and the rest of the code     myself.
* The conditional formatting that highlight positive change in green and negative change in red: I obtain the syntax from ChatGPT, which was different from how it was taught by the instructor.
* The adjustments to VBA script to enable it to run on every worksheet: inspired by instructor / he mentioned to add the "worksheet" function on all the "Cells".
* The retrieval of data, column creation, conditional formatting, values calculated, looping across worksheet and Github submission were successful.
