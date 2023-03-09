# VBA-challenge
VBA Challenge: Analyse Stock Data

Project is to create a VBA script to analyse multiple stock data over several years. It will go through a list of tickers and provide statistics such as yearly change, percentage change and total volume traded. It will also provide stock market statistics such as which ticker had the greatest yearly increase, greatest yearly decrease and also the traded the largest volume of stocks.
This VBA script can run through many worksheets at a click of the button to produce the statistics.
It was challenging to create a script that can run through 2,250,000  in a relatively short amount of time provide the statistics. Improving the run time of this scripts is the next step of improving this project. 

Table of Contents
- VBA File 1: This file includes 3 subroutines which will setup the tickers and the column names, then calculates the open and closing price and finally a subroutine to populate the yearly change statistics and the total volume column
- VBA File 2: This file will loop through the worksheets to provide the overall statistics for that year such as which ticker had the greatest percentage increase, greatest percentage decrease and the maximum volume per year
- 3 screenshots from the excel file

Various websites that helped improve the efficiency of the script:
Excel VBA loops - for each, for next, do while, nested & more (2022) Automate Excel. Available at: https://www.automateexcel.com/vba/loops/#do-until-loops (Accessed: March 9, 2023). 
Loop through worksheets / sheets with Excel VBA (no date) Excel Dashboards VBA. Available at: https://www.thesmallman.com/looping-through-worksheets (Accessed: March 9, 2023). 
PNRao (2022) VBA delete entire column Excel Macro Example code, Analysistabs. Available at: https://analysistabs.com/vba/delete-entire-column-excel-macro-example-code/ (Accessed: March 9, 2023). 
Melanie_ (2023) How to find the Max Value in a range defined with R1c1?, MrExcel Message Board. MrExcel Message Board. Available at: https://www.mrexcel.com/board/threads/how-to-find-the-max-value-in-a-range-defined-with-r1c1.1229358/ (Accessed: March 9, 2023).

