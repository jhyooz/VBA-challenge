# VBA-challenge
## Assignment
### Background
You are well on your way to becoming a programmer and Excel expert! In this homework assignment, you will use VBA scripting to analyze generated stock market data.

### Instructions
Create a script that loops through all the stocks for one year and outputs the following information:
* The ticker symbol 
* Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year. 
* The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year. 
* The total stock volume of the stock. The result should match the following image:
![Moderate solution](README%20Images/01%20-%20moderate_solution.jpg)
* Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:
![Hard solution](README%20Images/02%20-%20hard_solution.jpg)
* Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

> 📘 **NOTE** <br/>
> Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

### Other Considerations
* Use the sheet `alphabetical_testing.xlsx` while developing your code. This dataset is smaller and will allow you to test faster. Your code should run on this file in under 3 to 5 minutes. 
* Make sure that the script acts the same on every sheet. The joy of VBA is that it takes the tediousness out of repetitive tasks with the click of a button.

## Submission Notes
The vba script is in the `MultiYearStocks.bas` file however if that doesn't work for you, you can also see a copy of it in the `MultiYearStocks.txt` file.
1. The conditional formatting colors were taken from [Stack Overflow](https://stackoverflow.com/questions/27611260/what-are-the-rgb-codes-for-the-conditional-formatting-styles-in-excel)
2. I mostly used the credit card exercise we worked on in class, along with going over the lecture multiple times to get this figured out.
3. However I also used [Microsoft VBA documentation](https://learn.microsoft.com/en-us/office/vba/api/overview/) for info on how to format and other information (eg: FormatPercent, autofit, variable types - LongLong)
4. I also used [ExcelEasy](https://www.excel-easy.com/vba.html) for more information, examples and how to deal with VBA errors

## Submission 2 Notes 
Added conditional formatting for the Percent Change column (K).