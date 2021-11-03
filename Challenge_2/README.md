# Analyzing Stocks with Visual Basic for Applications

## Overview of Project
	
### Purpose
	
My client, Steve, would like some help analyzing stocks for his parents. They are interested in green energy alternatives and Daqo stock, which makes silicon wafers for solar panels. Steve would like to be able to easily determine how a stock has performed in a specific year. 


## Results

Visual Basic for Applications (VBA) was used to assist with analyzing stock data in Excel. A macro was created to be able to analyze how stocks performed in a specified year. The user can input the chosen year and learn the total daily volume and percent return that corresponds with each stock ticker. The percent return column has been formatted to make it easier for the reader to quickly determine if a stock has done well or not for a specified year. 

If we look at how the stocks performed in 2017, we see that they generally did well. Eleven stocks have positive returns and only one, TERP, has a negative return. 

[Insert green]

In, 2018, unfortunately, the stocks generally did not do well. Nearly every stock has a negative return except for ENPH and RUN. 

[Insert red]


Both stocks that had positive returns in 2018, were traded more often thus had higher total daily volumes than the year before.


[Insert circles around ENPH and RUN]


The refactored script takes fewer seconds to run than the original script. 


![VBA_slower_run_2017](https://user-images.githubusercontent.com/91852495/140187112-56b96d90-16cb-4c13-8fa9-e1b745a5a130.png)

![VBA_Challenge_2017](https://user-images.githubusercontent.com/91852495/140187356-6308134f-efef-4608-9854-c2c1a7fc5dea.png)


![VBA_slower_run_2018](https://user-images.githubusercontent.com/91852495/140187185-e0ffdeb4-637c-4b96-b293-595352c41414.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/91852495/140187303-51e9d2c3-f88e-4ebf-b7d2-90e060ef590e.png)


The reason for the difference can be the use of arrays in the refactored script rather than looping through individually.

[photos of array code versus regular loops]

