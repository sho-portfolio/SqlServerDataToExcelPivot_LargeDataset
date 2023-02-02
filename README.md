# SqlServerDataToExcelPivot_LargeDataset

If you have a huge sql server dataset (too big for Excel) and you want to pivot it here's what you can do:

Get the Data From SQL Server Into CSV

01.  In Query Analyzer
02.  Go to Tools -> Options -> Query Results -> SQL Server -> Results to Grid -> 
03.  Open a new query window (if you just run it on the existing one it wont show the column header names)!
04.  Select tickbox for 'Include column headers when copying or saving the results'
05.  Hit OK
06.  Run your query
07.  Right click in the grid results
08.  Click 'Save results As...'
09.  Save to desktop as csv file (desktop as at network drive may be super slow!)


Combine multiple CSV files (i.e. if you need to compare data from QA to Prod)

10.  Take all your csv files and append then in notepad (dont copy the header from each file, just 1)


Create the Pivot

11.  Open a Excel file
12.  Press ALT D P
13.  Select 'External data source'
14.  Click Next
15.  Click 'Get Data'
16.  Select '<New Data Source> and click OK
17.  Enter a name
18.  From the drop down list select Microsoft Access Text Driver (*.txt, *.csv)
19.  Select 'Connect'
20.  Untick 'Use Current Directory'
21.  Click Select Directory
22.  Select the folder where your file is
23.  Click OK (?)
24.  Click OK again
25.  From the Drop down box select the csv file
26.  Press OK
27.  Press OK again
28.  Select the file and then click the '>' button and click next
29.  Press Next again
30.  Press Next again
31.  Select 'Return Data to Microsoft Excel'
32.  Press Finish
33.  Press Next
34.  Press Finish
35.  Wait - it might take a minute
36.  And you have your pivot!!
