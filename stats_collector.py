import openpyxl

def gather_data(exported_data):
    wb = openpyxl.load_workbook(exported_data)

    sheet = wb.get_sheet_by_name('General')
    map_ID = sheet['B2'].value
    date = sheet['C2'].value
    date = date[:10]
    date = date.split('-')
    date.append(date.pop(0))
    date = '/'.join(date)
    map_name = sheet['F2'].value
    team1 = sheet['M2'].value
    team2 = sheet['N2'].value
    team1_score = sheet['O2'].value
    team2_score = sheet['P2'].value
    team1_score_first = float(sheet['Q2'].value)
    team2_score_first = float(sheet['R2'].value)
    team1_score_second = float(sheet['S2'].value)
    team2_score_second = float(sheet['T2'].value)
    winner = sheet['U2'].value
    win = ''
    side = ''
    sides = {team1: 'CT', team2: 'T'}

    #Translation for map names
    map_list = {'de_dust2': 'Dust II',
                'de_overpass': 'Overpass',
                'de_cbble': 'Cobblestone',
                'de_cache': 'Cache',
                'de_mirage': 'Mirage',
                'de_train': 'Train',
                'de_nuke': 'Nuke',
                'de_inferno': 'Inferno'}

    #List of SteamIDs for players whose stats we want to grab
    player_list = ['76561198049003674',  #Tim
                    '76561198055203865', #Ryan
                    '76561198049273178', #Collin
                    '76561198049028258', #Yip
                    '76561198044268933', #Sean
                    '76561198069453400', #Cal
                    '76561198098580186', #Sharfin
                    '76561198077398959', #Josh
                    '76561198156681310', #Jeph
                    '76561198070735886', #Tish
                    '76561198236096526'] #Harlan
    scores = {team1: team1_score, team2: team2_score}
    other_team = {team1:team2, team2:team1}

    sheet = wb.get_sheet_by_name('Players')
    player_stats = []
    ream_team_name = ''
    for num in range(2, 12):
        steam_ID = sheet['B' + str(num)].value
        if(steam_ID == '76561198055203865'      #Ryan
            or steam_ID == '76561198049273178'  #Collin
            or steam_ID == '76561198098580186'  #Sharfin
            or steam_ID == '76561198049003674'  #Tim
            or steam_ID == '76561198069453400'  #Cal
            or steam_ID == '76561198077398959'  #Josh
            or steam_ID == '76561198049028258'  #Yip
            or steam_ID == '76561198044268933'  #Sean
            or steam_ID == '76561198156681310'  #Jeph
            or steam_ID == '76561198070735886'  #Tish
            or steam_ID == '76561198236096526'  #Harlan
        ):
            ream_team_name = sheet['D' + str(num)].value                                   #Team Name
            ID = sheet['B' + str(num)].value                                               #Steam ID
            rating = sheet['R' + str(num)].value                                           #HLTV Rating
            ADR = sheet['W' + str(num)].value                                              #Average Damage Per Round
            HS = sheet['J' + str(num)].value                                               #Headshot Percentage
            kills = sheet['E' + str(num)].value                                            #Total Kills
            assists = sheet['F' + str(num)].value                                          #Total Assists
            deaths = sheet['G' + str(num)].value                                           #Total Deaths
            trades = sheet['AE' + str(num)].value
            KPR = str(kills/(scores[ream_team_name] + scores[other_team[ream_team_name]])) #Kills Per Round
            KDD = str(kills - deaths)                                                      #Kill Death Difference
            player = [ID, rating, trades, ADR, HS, KPR, KDD, kills, assists, deaths]
            player_stats.append(player)


    sheet = wb.get_sheet_by_name('Kills')
    for player in range(0,len(player_stats)):
        current_ID = player_stats[player][0]
        kills_half1 = 0
        kills_half2 = 0
        for num in range(2,152):
            if(sheet['E' + str(num)].value == ''):
                break
            elif(sheet['E' + str(num)].value == current_ID and sheet['P' + str(num)].value != ream_team_name):
                if(int(sheet['B' + str(num)].value) <= 15):
                    kills_half1 += 1
                else:
                    break
        kills_half2 = int(player_stats[player][7]) - kills_half1

        kills_half1 = kills_half1/(team1_score_first + team2_score_first)
        if(team1_score_first + team2_score_first < 15):
            kills_half2 = ''
        else:
            kills_half2 = kills_half2/(int(team1_score) + int(team2_score) - (team1_score_first + team2_score_first))
        if(sides[ream_team_name] == 'T'):
            player_stats[player].append(kills_half1)
            player_stats[player].append(kills_half2)
        else:
            player_stats[player].append(kills_half2)
            player_stats[player].append(kills_half1)

    if(winner == ream_team_name):
        win = 'W'
    elif(winner == other_team[ream_team_name]):
        win = 'L'
    else:
        win = 'D'
    side = sides[ream_team_name]

    map_stats = [side, date, win, map_list[map_name], scores[ream_team_name], scores[other_team[ream_team_name]]]

    row_array = []
    row_array += map_stats

    '''Our google sheet currently tracks 8 different players. These players are listed
        in player_list in the same order they show up on the google sheet.
        We loop through player_list, checking if this ID matches an ID we have in
        player_stats, indicating that player was playing in this specific match.
        If a player was not found to have been playing in this match we have to
        make sure that they're cells on the google sheet stay empty so we put empty
        strings where there stats would be. '''

    for num in range(0, 8):
        temp_ID = player_list[num]
        player_index = -1
        for num in range(0, len(player_stats)):
            if(player_stats[num][0] == temp_ID):
                player_index = num
                del player_stats[num][0]
                break
        if(player_index == -1):
            row_array += ['','','','','','','','','','','']
        else:
            row_array += player_stats[player_index]

    return row_array
