# Analyzing Stocks with Visual Basic for Applications

## Overview of Project
	
### Purpose
	
My client, Steve, would like some help analyzing stocks for his parents. They are interested in green energy alternatives and Daqo stock, which makes silicon wafers for solar panels. Steve would like to be able to easily determine how a stock has performed in a specific year. 




## Results

### Analysis of Stock Performance in 2017 & 2018

Visual Basic for Applications (VBA) was used to assist with analyzing stock data in Excel. A macro was created to be able to analyze how stocks performed in a specified year. The user can input the chosen year and learn the total daily volume and percent return that corresponds with each stock ticker. The percent return column has been formatted to make it easier for the reader to quickly determine if a stock has done well or not for a specified year. 

If we look at how the stocks performed in 2017, we see that they generally did well. Eleven stocks have positive returns and only one, TERP, has a negative return. 

![VBA_Challenge_Excel_2017](https://user-images.githubusercontent.com/91852495/140387461-b4e2f428-509b-43ac-995e-e566ed2a6519.png)


In, 2018, unfortunately, the stocks in general did poorly. Nearly every stock has a negative return except for ENPH and RUN. 

![VBA_Challenge_Excel_2018](https://user-images.githubusercontent.com/91852495/140388858-08027e5f-c783-40e2-b0de-728f0b916666.png)


Both stocks that had positive returns in 2018, were traded more often thus had higher total daily volumes than the year before.

![VBA_Challenge_Excel_2018 2](https://user-images.githubusercontent.com/91852495/140389952-d832f619-321f-4cb8-88e8-664a3de31869.png)

<br/>


### Comparing Script Speed
The refactored script takes fewer seconds to run than the original script. 

Below is the original script for 2017:

![VBA_slower_run_2017](https://user-images.githubusercontent.com/91852495/140187112-56b96d90-16cb-4c13-8fa9-e1b745a5a130.png)

Compare to the faster refactored script for 2017:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/91852495/140187356-6308134f-efef-4608-9854-c2c1a7fc5dea.png)


Below is the original script for 2018:

![VBA_slower_run_2018](https://user-images.githubusercontent.com/91852495/140187185-e0ffdeb4-637c-4b96-b293-595352c41414.png)

Compare to the faster refactored script for 2018:

![VBA_Challenge_2018](https://user-images.githubusercontent.com/91852495/140187303-51e9d2c3-f88e-4ebf-b7d2-90e060ef590e.png)

<br/>
The 3 additional arrays in the refactored script allows us to loop through the data one time thus taking less time to run. The original code takes longer since it must loop through individually to collect the same data.

Below shows the VBA code with 3 arrays included:

![VBA_refactored_script_a](https://user-images.githubusercontent.com/91852495/140386482-c67d7917-b800-4b45-9f71-645e6872873d.png)
![VBA_refactored_script_b](https://user-images.githubusercontent.com/91852495/140386569-cbc84c42-6984-4465-9a63-65cdcc419b35.png)
![VBA_refactored_script_c](https://user-images.githubusercontent.com/91852495/140386631-dedcbbd2-50bf-4699-93e9-0782d94b3021.png)

<br/>
Here is the original script that loops through individually:

![VBA_original_script](https://user-images.githubusercontent.com/91852495/140386305-939d7e16-3ce5-4dc4-93b6-c7e07e25dea3.png)


<br/>

## Summary

### Advantages/Disadvantages of Refactoring Code

The advantages of refactoring code include improving the readability of the code as well as making it more efficient. Code efficiency is important to reduce the time needed to run the code. However, the process of refactoring code may be time consuming which can be a disadvantage. 


### Pros and Cons of Refactoring the Original VBA Script

The pros of refactoring the original VBA script include streamlining the script and reducing the time it takes to run the script. The cons include increasing the complexity of the script with multiple arrays such as ticker(), tickerVolumes(), tickerStartingPrices(), and tickerEndingPrices()- as well as using tickerIndex. These factors could increase the likelihood of making a mistake and breaking the code. The efficiency and organization make up for the potential drawbacks. 
