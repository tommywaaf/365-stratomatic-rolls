import os
import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime

team_ids = [1734825, 1729423, 1727997, 1734739, 1734804, 1734657, 1734818, 1733907, 1734829, 1734807, 1734755, 1734492]
team_names = {
    1734825: 'SoCal Desert Eagle',
    1729423: 'Sherborn Colt',
    1727997: 'Framingham Chris',
    1734739: 'Chicago Sandbergs',
    1734804: 'Tartarian Giants',
    1734657: 'Cincy Buffalos',
    1734818: 'Boomtown Swagger',
    1733907: 'LafayetteSquare Cyclones',
    1734829: 'Portland Raiders',
    1734807: 'Reno Gamblers',
    1734755: 'San Diego Sweat Hogs',
    1734492: 'rosamond roadrunners'
}

headers = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
    "Accept-Encoding": "gzip, deflate, br",
    "Accept-Language": "en-US,en;q=0.9",
    "Connection": "keep-alive",
    "Cookie": "phpbb3_o4zc3_u=133119; phpbb3_o4zc3_k=; phpbb3_o4zc3_sid=4d731e7a165aecdf4b8884d8c7287e83; wp_token=0037f8bf86b3b739dcab949e2f0974dc; og_session=V8LM0HcP9SCQdN5OcXDyOGYdCriSw%2BMkVLqT5eFwIJSrak70l0hM2k5OZCwyF8z7EiPFxjB6fJDGzaeXCGPO6Cutw%2B3P3u0g0sD50NIFYiJgtrOTP%2BhS7%2BkikfNDfh%2Fe8KJK%2FmaNnmUOxrVAD6LdaCXWaxSv6NngPqWmyEJRrNApW06awqMCOC5E7JF8OUUJ0ch6elTT0WNJ46qDjm5R1KV8VJhfw4ncfJOYtuXiku7a6FWw61evCzThezIhhTfX22HMQiNSWX1JI%2B%2FdgvcZwJw%2FILiIBOrixx%2BcdE1XF2bqFdx2%2B7At%2Fm%2Fs%2FvZhv6FZ4dU7WG5BqRtIzSuojI%2FRlXJCni1E5PNUx30FNwVzOEUrkczKacaavzemxrvlbiz0HLGmaifiSkZxC8LBXonoBzyN5f6ubKjt4WoPddjU92oj4BeOhsTsRqUcjLfXFCP626hq3ENoYaY%2BHz22sxJZeeHp2hGIh%2Ftv94X5LI2Yu%2BDkSX0aDx3z1qHlA5cKd3nOewKdtPd1gby%2FMmBPLL04KUbsvVUmF3x9np1L0ga5V5pqrpgbwuiA0r7IaXfnSGcHfzzzqg%2FD%2BY482MEBIa1nm23g17IpdQCEsOWx4t8W3zbo61%2F1YFo7Qtmf%2Fi1edBik1O6wSbT06SYi2V2p%2F0NPIMDdi2%2FOSCTSSn%2FdgYq9y%2B1FXxL4MtVDXZPaufK5WChZ8eyn4ogXv%2BsYeiQ1fR8PKemRvNZQCUgaVLDIqtv%2BUCwCdZC1w10OAoaUMMYKf5jizkoX1JkylHNafibENjX6Dt9SOP1Zg%2Fhjb2EPtOstMM4%2BaFg46Wi%2FiQM333QUcdLDkS1N3IV3KSmYL7MMYDQ8buKIyIfgOn2OzBGKIBKKFBXB6Arslq3RuXZM7ZbwxNnj; AWSALB=RwJ75nzXxnKILXxlrt8SsoeVzhaD6VKBPpiuNt4M6UNZKxeyi1I+EIX4cSNKb3EmNQHaoloVxBst8g6mPfwr+kMO8oLf9zSXKs3wVBa5+XDSLNSKRIKvvDmKddiH; AWSALBCORS=RwJ75nzXxnKILXxlrt8SsoeVzhaD6VKBPpiuNt4M6UNZKxeyi1I+EIX4cSNKb3EmNQHaoloVxBst8g6mPfwr+kMO8oLf9zSXKs3wVBa5+XDSLNSKRIKvvDmKddiH",
    "Host": "365.strat-o-matic.com",
    "Referer": "https://365.strat-o-matic.com/team/1728243",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "same-origin",
    "Sec-Fetch-User": "?1",
    "Upgrade-Insecure-Requests": "1",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36",
    "sec-ch-ua": "Google Chrome;v=111, Not(A:Brand;v=8, Chromium;v=111",
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": "Windows",
}

# Filename of your Excel file
filename = 'teams_data.xlsx'

# Create a list to store all teams data
all_teams_data = []

# Iterate over each team_id
for team_id in team_ids:
    url = f"https://365.strat-o-matic.com/team/misc/{team_id}"
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.text, 'html.parser')
    team_name = team_names[team_id]

    # Now, do the extraction process inside the loop
    for row in soup.find_all('tr', class_=['odd', 'even']):
        player_name_cell = row.find('td', {'name': 'name'})
        if player_name_cell:
            player_name = player_name_cell.get_text().strip()
            position_cell = row.find('td', {'name': 'pos'})

            if position_cell:  # If the player is a hitter
                hitter_rolls = row.find('td', {'name': 'hit'}).get_text()
                pitcher_rolls = row.find('td', {'name': 'pit'}).get_text()
            else:  # If the player is a pitcher
                hitter_rolls = row.find('td', {'name': 'pit'}).get_text()
                pitcher_rolls = row.find('td', {'name': 'hit'}).get_text()

            # Append the data to the list
            all_teams_data.append({
                'Team': team_name,
                'Player Name': player_name,
                'Hitter Rolls': hitter_rolls,
                'Pitcher Rolls': pitcher_rolls
            })


# Convert the list to a DataFrame
df = pd.DataFrame(all_teams_data)

# Check if the Excel file exists
if not os.path.isfile(filename):
    # If the file doesn't exist, create it
    df.to_excel(filename, sheet_name=datetime.now().strftime('%Y-%m-%d'), index=False)
else:
    # If the file exists, append the DataFrame to a new sheet
    with pd.ExcelWriter(filename, engine='openpyxl', mode='a') as writer:
        df.to_excel(writer, sheet_name=datetime.now().strftime('%Y-%m-%d'), index=False)
