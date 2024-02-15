# stock_market_analysis #
This VBA code is used to summarize the "Yearly Change", "Percent Change", and "Total Volume" per stock ticker. Once each ticker in the spreadsheet has been summarized, the code will find the "Greatest % Increase", "Greatest % Decrease", and "Greatest Total Volume" and output the results in another summary table on the same sheet for easier analysis. 

## How does the code work?

1. Dim ws as Worksheet to set up a loop so that the code runs through all given worksheets in the workbook
2. Inserts headers for respective columns for both "Summary Tables"
3. Dimming neccesary variables
4. Assigns variables "j", "total", "change", and "start" intial values
5. Assings the for loop for 2 to "FinalRow_A" which finds the last populated row in Column A
6. Starts the loop through rows to find the first row in Column A that is not equal to the previous row to identify the switch between different tickers
7. If a new ticker, then it stores total and if total = 0. Current ticker is outputted and "YC", "PC", "TV" are set to 0
8. If not it will find the first non-zero value in Column K from "start"
9. Calculates "YC" and "PC" and then updates "start" value to next row up
10. Outputs "Ticker", "YP", "PC", and "TV" to respective cells
11. Conditional formatting of cells for "Yearly Change" and "Percent Change" - Cells that have a positive value (>0) are colored green and cells that have a negative value (<0) are colored red using the "Select Case" function
12. Then code resets "total" to 0 and "change" to 0 and moves the outputs down a cell with (j=j+1)
13  Finds the "Greatest % Increase" and "Greatest % Decrease" in Column K and outputs the results of the "Greatest % Increase" and "Greatest % Decrease" and matches the repsective ticker for each of those results in that row
15. Then it finds the "Greatest Total Volume" in Column L and outputs the result of the "Greatest Total Volume" with the corresponding ticker in that row
16. Lastly the code ends with a "Next" that allows the functionality of looping through each sheet of the workbook
