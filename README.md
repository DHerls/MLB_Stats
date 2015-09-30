# MLB_Stats
Written in python.

1. Scrapes statistics from http://mlb.mlb.com/stats/.
2. Generates and Excel workbook with 40 random players with more than 20 at bats and who aren't pitchers.  
Only puts Team, Position, Games Played, Strike Outs, and Batting Averages.
3. Creates a separate sheet which contains frequency distribution charts for numerical player data

**NOTE:** Excel for some reason does not like all of the formulas used.  In order to fix the #NAME error that appears in 
the Frequency sheet, the user must manually enter cells J2 and J12 to refresh the formulas.
