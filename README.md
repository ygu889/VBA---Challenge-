# VBA---Challenge- 
VBA scripting to analyze generated stock market data 
Screeshots of the analysis results
READMe
I started the code with using class example - credit_charges solution for the loops code example which helped me with creating the code to loop and find the ticker symbol and the Total volume of each ticker

For the percentage change: I looked up online: how to set the open price for each ticker to not change as I loop through the rows: my initial code was open price = cells (i,3).Value and loop reset open price = cells (i+1,3).Value which resulted in taking the price difference of the opening and closing on the last day of the year rather than the open price at the beginning of the year and the last day of the year. I change open price = cells (2,3).Value

For the greatest summary table: For determining the greatest % increase and decrease i had to look up how to get the value once I found the maximum and minimum and max volume to print in the table created to store: ws.Range("Q2").Value = ws.Range("K" & i).Value
ws.Range("P2").Value = ws.Range("I" & i).Value
