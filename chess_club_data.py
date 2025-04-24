import requests
import pandas as pd
from collections import Counter, defaultdict
from datetime import datetime
import numpy as np
import time

CLUB_NAME = "torneios-cursos-gm-krikor"
HEADERS = {'User-Agent': 'your_username@example.com'}
BASE_URL = "https://api.chess.com/pub"
YEARS = [2024, 2025]
GAME_CLASSES = ['blitz', 'rapid']

# Helper functions
def get_quarter(month):
    return (month - 1) // 3 + 1

def get_current_ratings(username):
    url = f"{BASE_URL}/player/{username}/stats"
    response = requests.get(url, headers=HEADERS)
    if response.status_code != 200:
        return {}
    data = response.json()
    ratings = {}
    for time_class in GAME_CLASSES:
        key = f"chess_{time_class}"
        if key in data and 'last' in data[key]:
            ratings[time_class] = data[key]['last']['rating']
    return ratings

def process_member(username):
    total_games = 0
    white_openings = Counter()
    black_defenses = Counter()
    quarterly_ratings = defaultdict(lambda: defaultdict(list))
    first_ratings = {tc: None for tc in GAME_CLASSES}

    current_ratings = get_current_ratings(username)

    for year in YEARS:
        for month in range(1, 13):
            
            url = f"{BASE_URL}/player/{username}/games/{year}/{month:02d}"
            response = requests.get(url, headers=HEADERS)
            if response.status_code != 200:
                #print(f"[WARNING] No games found for {year}-{month:02d}.")
                continue
            games = response.json().get('games', [])
            
            total_games += len(games)
            for game in games:
                time_class = game.get('time_class', None)
                if time_class not in GAME_CLASSES:
                    continue
                end_time = game.get('end_time', None)
                date_obj = datetime.utcfromtimestamp(end_time) if end_time else None
                quarter_key = f"{year}Q{get_quarter(date_obj.month)}" if date_obj else None
                
                # Openings/defenses collection
                pgn = game.get('pgn', '')
                eco = None
                eco_url = None
                for line in pgn.split('\n'):
                    if line.startswith('[ECO '):
                        eco = line.split('"')[1]
                    if line.startswith('[ECOUrl '):
                        eco_url = line.split('"')[1]

                if eco_url:
                    name = eco_url.replace("https://www.chess.com/openings/", "").replace("-", " ")
                else:
                    name = eco  # fallback to ECO code if URL missing

                if game['white']['username'].lower() == username.lower():
                    player = game['white']
                    if name:
                        white_openings[name] += 1
                else:
                    player = game['black']
                    if name:
                        black_defenses[name] += 1


                # Ratings
                rating = player.get('rating', None)
                if rating and quarter_key:
                    quarterly_ratings[quarter_key][time_class].append(rating)
                    if not first_ratings[time_class]:
                        first_ratings[time_class] = rating

    # Compile member data
    preferred_opening = white_openings.most_common(1)[0] if white_openings else ('None', 0)
    preferred_defense = black_defenses.most_common(1)[0] if black_defenses else ('None', 0)
    member_data = {
        'username': username,
        'total_games_played': total_games,
        'current_blitz_rating': current_ratings.get('blitz', None),
        'current_rapid_rating': current_ratings.get('rapid', None),
        'first_blitz_rating': first_ratings['blitz'],
        'first_rapid_rating': first_ratings['rapid'],
        'preferred_opening_white': preferred_opening[0],
        'preferred_opening_white_count': preferred_opening[1] if white_openings else 0,
        'preferred_defense_black': preferred_defense[0],
        'preferred_defense_black_count': preferred_defense[1] if black_defenses else 0
    }

    # Add quarterly ratings and volatilities
    for quarter in sorted(quarterly_ratings.keys()):
        for tc in GAME_CLASSES:
            ratings = quarterly_ratings[quarter].get(tc, [])
            if ratings:
                member_data[f'final_rating_{tc}_{quarter}'] = ratings[-1]
                member_data[f'volatility_{tc}_{quarter}'] = np.std(ratings)
            else:
                member_data[f'final_rating_{tc}_{quarter}'] = None
                member_data[f'volatility_{tc}_{quarter}'] = None

    # Calculate increments
    for tc in GAME_CLASSES:
        try:
            increment = member_data[f'current_{tc}_rating'] - member_data[f'first_{tc}_rating']
            member_data[f'increment_{tc}'] = increment if increment is not None else 0
        except:
            member_data[f'increment_{tc}'] = 0

    return member_data

# Main script
from openpyxl import load_workbook
start_time = time.time()
club_url = f"{BASE_URL}/club/{CLUB_NAME}/members"
response = requests.get(club_url, headers=HEADERS)
members = [member['username'] for category in response.json().values() for member in category]

member_stats = []
members_df = pd.DataFrame(member_stats)
# Ensure all expected columns are present consistently
columns_order = [
  'username', 'total_games_played',
  'current_blitz_rating', 'current_rapid_rating',
  'first_blitz_rating', 'first_rapid_rating',
  'increment_blitz', 'increment_rapid',
  'preferred_opening_white', 'preferred_opening_white_count',
  'preferred_defense_black', 'preferred_defense_black_count'
] + [col for col in members_df.columns if col.startswith('final_rating_') or col.startswith('volatility_')]
members_df = members_df.reindex(columns=columns_order)
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
output_file = f'club_stats_{timestamp}.xlsx'

# Initialize the Excel file
with pd.ExcelWriter(output_file, mode='w', engine='openpyxl') as writer:
    members_df.to_excel(writer, sheet_name='Members', index=False)
for idx, username in enumerate(members):
    member_start = time.time()
    print(f"[INFO] Processing {idx+1}/{len(members)}: {username}")
    try:
        member_data = process_member(username)
        member_stats.append(member_data)
        members_df = pd.DataFrame(member_stats)
        with pd.ExcelWriter(output_file, mode='w', engine='openpyxl') as writer:
            members_df.to_excel(writer, sheet_name='Members', index=False)
        elapsed = time.time() - member_start
        avg_time = (time.time() - start_time) / (idx + 1)
        est_remaining = avg_time * (len(members) - idx - 1)
        print(f"[INFO] Done {username}. Total games: {member_data['total_games_played']}. ETA: {est_remaining/60:.1f} mins.")
    except Exception as e:
        print(f"[WARNING] Failed for {username}: {e}")

# DataFrames
members_df = pd.DataFrame(member_stats)

# Club summary
top_active = members_df[['username', 'total_games_played']].sort_values(by='total_games_played', ascending=False).head(10)

# Total points (rating sums)
def sum_points(df, quarters, tc):
    cols = [f'final_rating_{tc}_{q}' for q in quarters]
    return df[cols].sum().sum()

H1_2024_quarters = ['2024Q1', '2024Q2']
H2_2024_quarters = ['2024Q3', '2024Q4']
Q1_2025_quarters = ['2025Q1']

summary_data = {
    'top_10_most_active': top_active.values.tolist(),
    'total_points_H1_2024': sum_points(members_df, H1_2024_quarters, 'blitz') + sum_points(members_df, H1_2024_quarters, 'rapid'),
    'total_points_H2_2024': sum_points(members_df, H2_2024_quarters, 'blitz') + sum_points(members_df, H2_2024_quarters, 'rapid'),
    'total_points_Q1_2025': sum_points(members_df, Q1_2025_quarters, 'blitz') + sum_points(members_df, Q1_2025_quarters, 'rapid')
}

summary_df = pd.DataFrame([summary_data])

# Save to Excel
with pd.ExcelWriter(output_file, mode='a', engine='openpyxl') as writer:
    summary_df.to_excel(writer, sheet_name='Summary', index=False)

total_time = time.time() - start_time
print(f"[INFO] Club data collection complete in {total_time/60:.2f} minutes. Saved to 'club_stats.xlsx'")
