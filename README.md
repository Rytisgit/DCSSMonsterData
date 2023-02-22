# CSV of DCSS Monster Data and Weapon Damage Calculation

### A csv file parser to look into monster data easily, along with a spreadsheet to calculate a comparison between 2 weapons taking into account monster ac.

![image](https://user-images.githubusercontent.com/36188103/220609936-ea0d1115-1878-44dd-ba3f-a68947988704.png)




## The regex I had a very hard time making to update the csv
`,\s"([^"]*)"[\s\S]*?(?=.+?(.+?(\d+), (\d+),))(.+?(\d+), (\d+),)\s(.+?(\d+), (\d+))\s*`

used in https://regexr.com/

remove comments in mon-data.h with these 2 regexes and replace with empty using regex replace for perfect csv:

`//.*`

`^(\s)*$\n`

In List function used the following capture groups to format csv data:

`$1,$6,$7,$9,$10\n`

Then added headers to data :

`name,hd,datahp,ac,ev`

## The Spreadsheet

Available at: https://docs.google.com/spreadsheets/d/1tiYxdtwKVlAN2hdXsOYLV0STmu-uNI4PeCaAZdLpih8/edit?usp=sharing

Type in desired weapon and character stats, press update button to see new ratios(need to authorise the script for the first time). The script is included for viewing in the GoogleSheetsExport folder.

Big Thanks to powerbf, whose implementation of damage calculation I heavily based the google sheets app script on. https://github.com/powerbf/crawl-helper/ / https://powerbf.github.io/crawl-helper/
