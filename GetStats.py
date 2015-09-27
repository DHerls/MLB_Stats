from urllib.request import urlopen
import json
import random
import openpyxl
import csv


URLBase = "http://mlb.mlb.com/pubajax/wf/flow/stats.splayer"
URLSeason = "?season=2015"
URLOrder = "&sort_order=%27desc%27"
URLColumn = "&sort_column=%27avg%27"
URLType = "&stat_type=hitting"
URLPage = "&page_type=SortablePlayer"
URLGame = "&game_type=%27R%27"
URLPlayer = "&player_pool=ALL"
URLSeasonType = "&season_type=ANY"
URLSportCode = "&sport_code=%27mlb%27"
URLResults = "&results=1000"
URLPageNum = "&recSP=1"
# Get Every Player on one page
URLResultNum = "&recPP=1250"


def getURL():
    return URLBase + URLSeason + URLOrder + URLColumn + URLType + URLPage + URLGame + URLSeasonType + URLSportCode + URLResults + \
           URLPageNum + URLResultNum


siteData = json.loads(urlopen(getURL()).read().decode("UTF-8"))
# Gets list of player data from raw site data
playerList = siteData["stats_sortable_player"]["queryResults"]["row"]


listToDelete = []

# If player has <20 at bats or is a pitcher, they are marked for removal
for player in playerList:
    if int(player.get("ab")) < 20:
        listToDelete.append(player)
    else:
        if player.get("pos") == "P":
            listToDelete.append(player)

# The players marked for removal are removed
for player in listToDelete:
    playerList.remove(player)

length = len(playerList)

randomNumbers = []

# Generates a list of 40 random numbers
while len(randomNumbers) < 40:
    rand = random.randint(0, length-1)
    if not randomNumbers.__contains__(rand):
        randomNumbers.append(rand)

randomPlayers = []

# Uses the random numbers to retrieve the corresponding players
for num in randomNumbers:
    randomPlayers.append(playerList[num])

# Create Excel Workbook
wb = openpyxl.Workbook()
# Get Worksheet
dataSheet = wb.get_active_sheet()
# Set Worksheet title to Data
dataSheet.title = "Data"

dataList = ["name_display_first_last", "team_abbrev", "pos", "g", "so", "avg"]

# Put titles for the columns
dataSheet['A1'] = "Name"
dataSheet['B1'] = "Team"
dataSheet['C1'] = "Position"
dataSheet['D1'] = "Games"
dataSheet['E1'] = "Strikeouts"
dataSheet['F1'] = "Average"

# Put data in the Excel workbook
row = 2
for player in randomPlayers:
    column = 'A'
    for key in dataList:
        value = player.get(key)
        # See if the value is an integer
        try:
            value = int(value)
        except ValueError:
            # See if the value is a decimal
            try:
                value = float(value)
            except ValueError:
                False

        dataSheet[column + str(row)] = value
        column = chr(ord(column)+1)
    row += 1

frequencySheet = wb.create_sheet()
frequencySheet.title = 'Frequency'

for key, val in csv.reader(open("Frequency.csv")):
    frequencySheet[key] = val

frequencySheet.merge_cells("A1:G1")
frequencySheet.merge_cells("A11:G11")
frequencySheet.merge_cells("A21:G21")



wb.save("test.xlsx")
