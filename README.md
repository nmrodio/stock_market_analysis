# stock_market_analysis #
This VBA code is used to summarize the "Yearly Change", "Percent Change", and "Total Volume" per stock ticker. Once the each ticker in the spreadsheet has been summarized, the code will find the "Greatest % Increase", "Greatest % Decrease", and "Greatest Total Volume" and output the results in another summary table on the same sheet for easier analysis. 

## How does the code work?

1. Dim ws as Worksheet to set up a loop so that the code runs through all sheets given in the workbook
2. Inserts headers for respective columns
3. Dimming neccesary variables
4. Assigns variables intial values 
5. Loops through rows to find the first row that is NOT EQUAL to the previous row and then stores respective variables using an if statement
6. Then if total = 0 the code will assign new values to respective cells for summary but if not the code will loop through the rows and find the first non-zero starting value and updates the "start" value
7. Calculates "Yearly Change" and "Percent Change" and then starts looking for the next ticker until all ticker options are exhausted and starts to output summarized results
8. Conditional formatting of cells for "Yearly Change" and "Percent Change" - Cells that have a positive value (>0) are colored green and cells that have a negative value are colored red using the "Select Case" function
9. Then the variables are reset for the loop to allow for correct calculations per ticker
10. Then total volume is calculated (total was already assigned to a dynamic range reference above for output of results per ticker)
11. Finally the last chunk of code finds the "Greatest % Increase" and "Greatest % Decrease" in Column K and outputs the results of the "Greatest % Increase" and "Greatest % Decrease" and matches the repsective ticker for each of those results in that row
12. Then it finds the "Greatest Total Volume" in Column L and outputs the result of the "Greatest Total Volume" with the corresponding ticker in that row
13. Lastly the code ends with a "Next" that allows the functionallity of looping through each sheet of the workbook
