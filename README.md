# CSV of DCSS Monster Data
This is a simple csv file to look into monster data easily
## The regex I had a very hard time making to update it
`,\s"([^"]*)"[\s\S]*?(?=.+?(.+?(\d+), (\d+),))(.+?(\d+), (\d+),)\s(.+?(\d+), (\d+))\s*`

used in https://regexr.com/

remove comments in mon-data.h with these 2 regexes and replace with empty for perfect csv:
`//.*`
`^(\s)*$\n`

In List function used the following capture groups: `$1,$6,$7,$9,$10\n`

Add headers to data :
`name,hd,datahp,ac,ev`

google sheet available: https://docs.google.com/spreadsheets/d/1tiYxdtwKVlAN2hdXsOYLV0STmu-uNI4PeCaAZdLpih8/edit?usp=sharing
