import pandas as pd

excel_pos = 'Modern_Draft_Position_All.xlsx'
position = pd.read_excel(excel_pos)
excel_stat = 'Modern_Draft_Stats.xlsx'
allStats = pd.read_excel(excel_stat,'Data_of_Interest')
stats = allStats.drop_duplicates(subset= ['Player','Year'])
players = pd.read_excel(excel_stat,'People_of_Interest')

    
def playerPrime(a,b):
#a is a player, this function returns his 5 best seasons based on statistic b
    playersStats = stats.loc[stats['Player']==a]
    seasons = playersStats.nlargest(5,b)
    if seasons.shape[0] == 0:
        total = 0
        positionRow = players.loc[players['Name']==a]['Pick']
        draftPosition = positionRow.iloc[0]
        classRow = players.loc[players['Name']==a]['Draft Class']
        draftClass = classRow.iloc[0]
    else:
        total = seasons[b].sum()
        draftClass = seasons.iat[0,53]
        draftPosition = seasons.iat[0,54]
    prime = pd.DataFrame ({'Player': [a], 'Draft Class': [draftClass] ,'Draft Position': [draftPosition], b: [total]})
    return(prime)
        
    
def totPrimeFile(b):
#b is the statistic
    df = pd.DataFrame()
    allPlayer = []
    
    for player in players['Name']:
        res = playerPrime(player,b)
        allPlayer.append(res)
        
    df = pd.concat(allPlayer)
    df.to_excel("Prime_" + b + ".xlsx", index=False)
    
    

