import csv
import os
import pandas as pd
import xlsxwriter


## TODO: this should probably go somewhere else
GOOD_COLOR = "#00FF00"
NEUTRAL_COLOR = "#FFFFFF"
BAD_COLOR = "#FF6347"

# TODO: need something better for when names differ slightly
#  between projections and draft platform rankings
def fix_name(name):
    return name.replace("Jr.", "").replace("Sr.", "").replace("III", "").replace("II", "").strip()


class PPR:
    weights = {
        "PASS_YDS": 0.04,
        "PASS_TDS": 4,
        "PASS_INTS": -2,
        "RUSH_YDS": 0.1,
        "RUSH_TDS": 6,
        "REC_YDS": 0.1,
        "REC": 1,
        "REC_TDS": 6,
        "FL": -2,
        # All stats appearing in data source should be accounted for, for completeness
        "PASS_ATT": 0,
        "PASS_CMP": 0,
        "RUSH_ATT": 0
    }

    @staticmethod
    def score(stat_dict):
        if stat_dict == {}:
            return None
        score = 0
        for stat in stat_dict:
            score += stat_dict[stat] * PPR.weights[stat]
        # Divide by 17 to get per game score as it is easier to visualize
        return score / 17


class Player:
    def __init__(self, name, position, team, platform_ranking):
        self.name = name
        self.position = position
        self.team = team
        self.overall_rank = platform_ranking[0]
        self.position_rank = platform_ranking[1]
        self.average_projection = {}
        self.high_projection = {}
        self.low_projection = {}

    def get_score(self, scoring=PPR):
        return (scoring.score(self.low_projection),
                scoring.score(self.average_projection),
                scoring.score(self.high_projection))

    def set_projections(self, projection_type, data):
        if projection_type == 'average':
            self.average_projection = data
        elif projection_type == 'high':
            self.high_projection = data
        elif projection_type == 'low':
            self.low_projection = data

    def __repr__(self):
        return f"Player({self.position} {self.name})"


def get_baseline_projections(player, baseline_player):
    # Returns projections relative to the baseline player, plus the range
    baseline_avg = baseline_player.get_score()[1]
    low, avg, high = player.get_score()
    low_relative = low - baseline_avg
    avg_relative = avg - baseline_avg
    high_relative = high - baseline_avg
    return low_relative, avg_relative, high_relative, high - low


def parse_csv(file_path, position, platform_rankings):
    players = {}
    current_player = None
    current_type = None

    with open(file_path, mode='r') as file:
        reader = csv.reader(file)
        headers = next(reader)  # Read the header

        # First row after headers is garbage
        next(reader)

        # Determine the columns to include (ignore "FPTS")
        relevant_columns = [i for i, header in enumerate(headers) if header not in ("FPTS")]

        for row in reader:
            # Some rows at the bottom are empty
            if len(row) == 1:
                continue
            if row[0].strip():  # New player entry
                player_name = row[0].strip()
                team = row[1].replace("high", "")
                if player_name not in platform_rankings:
                    if (fixed_name := fix_name(player_name)) in platform_rankings:
                        platform_rankings[player_name] = platform_rankings[fixed_name]
                        del platform_rankings[fixed_name]
                    else:
                        if len(platform_rankings) > 0:
                            print(f"Error: {player_name} not found in platform rankings")
                        platform_rankings[player_name] = (-1, -1)
                current_player = Player(player_name, position, team, platform_rankings[player_name])
                current_type = 'average'
                players[player_name] = current_player

            elif 'high' in row[1]:
                current_type = 'high'
            elif 'low' in row[1]:
                current_type = 'low'

            # Parse stats into a dictionary, skipping "Team" and "FPTS"
            stats = {}
            for i in relevant_columns[2:]:  # Skip the first two relevant columns (Team, Player)
                if row[i].strip():
                    stats[headers[i]] = float(row[i].replace(',',''))

            current_player.set_projections(current_type, stats)

    return players


def parse_platform_rankings(file_path):
    # Process and store rankings from the draft platform (for comparison)
    platform_rankings = {}  # player : (overall_rank, position_rank)
    with open(file_path, mode='r', encoding='utf-8-sig') as file:
        reader = csv.reader(file)
        headers = next(reader)
        assert len(headers) == 3, "Draft platform rankings file should only have 3 columns"

        # There should be 3 columns: one "Name", one "Position", and the other with the name of the platform
        name_col = headers.index("Name")
        position_col = headers.index("Position")
        rank_col = set(range(len(headers))).difference({name_col, position_col}).pop()
        platform_name = headers[rank_col]

        # keep counters for each position to determine position rank
        pos_rank = {}
        for i, row in enumerate(reader):
            if (position := row[position_col]) not in pos_rank:
                pos_rank[position] = 1
            platform_rankings[row[name_col]] = (i + 1, pos_rank[position])
            pos_rank[position] += 1
    return platform_name, platform_rankings


def parse_team_rankings(file_path):
    try:
        f = open(file_path, "r", encoding="latin-1")
    except:
        return {}
    else:
        with  f:
            reader = csv.reader(f)
            next(reader)  # skip headers

            team_rankings = {}
            for i, row in enumerate(reader):
                team_rankings[row[0]] = i + 1

        return team_rankings


if __name__ == '__main__':
    # Process and store rankings from the draft platform (for comparison)
    ranking_data_dir = "draft_platform_rankings"
    platform_name = None
    platform_rankings = {}
    for f in os.listdir(ranking_data_dir):
        if not f.endswith(".csv"):
            continue
        platform_name, platform_rankings = parse_platform_rankings(os.path.join(ranking_data_dir, f))

    # Gather projections from Fantasy Pros
    data_dir = "fp_data"
    players = {"QB": {}, "TE": {}, "RB": {}, "WR": {}, "Overall": {}}
    for f in os.listdir(data_dir):
        if not f.endswith(".csv"):
            continue
        position = f.split(".")[0].split("_")[-1]
        player_dict = parse_csv(os.path.join(data_dir, f), position, platform_rankings)
        players[position] = player_dict
        players["Overall"].update(player_dict)

    # Baseline is the number of that position drafted through 9 round last year plus 1
    baseline = {
        "QB": 11,
        "TE": 12,
        "RB": 39,
        "WR": 49
    }
    baseline["Overall"] = sum(baseline.values())
    # Only keep a certain amount of player projections per position
    # Rule of thumb here will be double the amount needed for rosters
    position_limits = {
        "QB": 12 * 2,
        "TE": 12 * 2,
        "RB": 36 * 2,
        "WR": 48 * 2
    }
    position_limits["Overall"] = sum(position_limits.values())

    # Find projection of baseline player at every position
    # This is what we will compare against
    data = {}  # player_name : (pos, pos_low_proj, pos_avg_proj, pos_high_proj)
    baseline_players = {}
    # Loop through positions in reverse order so that combined position categories process last
    for position in sorted(baseline.keys(), reverse=True):
        if position in ["QB", "TE", "RB", "WR"]:
            # Process a single position
            # Set headers
            rows = [f"Name,Team,Low,Avg,High,Range,{platform_name}\n"]

            # Sort players by their projected score
            sorted_players = sorted(players[position].values(), key=lambda P: P.get_score()[1], reverse=True)
            # Save baseline player for this position
            baseline_player = sorted_players[baseline[position]]
            baseline_players[position] = baseline_player

            # Generate one row per player, up to the positional limit for players shown
            for rank, player in enumerate(sorted_players[:position_limits[position]]):
                low_relative, avg_relative, high_relative, range_ = get_baseline_projections(player, baseline_player)
                # Calculate difference between projected rank and rank on drafting platform
                # If rank data is missing, make rank_diff very small, so it's obvious on the sheet
                rank_diff = -9999 if player.position_rank == -1 else player.position_rank - (rank + 1)
                rows.append(','.join([player.name, player.team, f"{low_relative:.1f}", f"{avg_relative:.1f}",
                                      f"{high_relative:.1f}", f"{range_:.1f}", f"{rank_diff}\n"]))
                # Save this data to use when processing a combined position sheet
                data[player.name] = (low_relative, avg_relative, high_relative, range_)

        else:
            # Process a combination of positions
            # Set headers
            rows = [f"Name,Pos.,Team,Low,Avg,High,Range,{platform_name}\n"]

            # Collect all positional scores for the players
            players_with_scores = []
            for player in players[position].values():
                if player.name in data:
                    players_with_scores.append((player, data[player.name]))
                else:
                    info = get_baseline_projections(player, baseline_players[player.position])
                    data[player.name] = info
                    players_with_scores.append((player, info))

            # Sort players by their avg relative projected score
            # This will create a ranking for positional scarcity
            sorted_players = sorted(players_with_scores, key=lambda t: t[1][1], reverse=True)

            # Generate one row per player, up to the positional limit for players shown
            for rank, player_tuple in enumerate(sorted_players[:position_limits[position]]):
                # Unpack info from tuples
                player, player_info = player_tuple
                low_relative, avg_relative, high_relative, range_ = player_info
                # Calculate difference between projected rank and rank on drafting platform
                # If rank data is missing, make rank_diff very small, so it's obvious on the sheet
                rank_diff = -9999 if player.overall_rank == -1 else player.overall_rank - (rank + 1)
                # Add row
                rows.append(','.join([player.name, player.position, player.team,
                                      f"{low_relative:.1f}", f"{avg_relative:.1f}",
                                      f"{high_relative:.1f}", f"{range_:.1f}", f"{rank_diff}\n"]))

        # Write results to CSV
        with open(os.path.join("output", f"{position}.csv"), "w") as f:
            f.writelines(rows)

    # parse team rankings
    team_rankings = parse_team_rankings(os.path.join("notes", "teams.csv"))

    # Create a new Excel workbook
    workbook = xlsxwriter.Workbook("DraftSheet.xlsm")

    # Define cell formats
    base_format = workbook.add_format({
        'font_size': 14
    })
    round_end_format = workbook.add_format({
        'font_size': 14,
        'bottom': 2  # thick bottom border
    })
    undrafted_format = workbook.add_format({
        'bold': True,
        'font_color': 'black',
    })
    drafted_format = workbook.add_format({
        'bold': False,
        'font_color': 'gray',
        'font_strikeout': True
    })

    # Define formats for each position with new colors
    qb_format = workbook.add_format({'bg_color': '#FFD700'})  # Gold for QB
    rb_format = workbook.add_format({'bg_color': '#66CCFF'})  # Blue for RB
    wr_format = workbook.add_format({'bg_color': '#CC99FF'})  # Purple for WR
    te_format = workbook.add_format({'bg_color': '#FF9900'})  # Orange for TE

    for file_name in sorted(os.listdir("output")):
        if not file_name.endswith(".csv"):
            continue

        # Load each CSV into a pandas DataFrame
        df = pd.read_csv(os.path.join("output", file_name))

        # Add a new sheet to the workbook
        worksheet_name = file_name.split(".")[0]
        worksheet = workbook.add_worksheet(worksheet_name)

        # Write headers and data to the worksheet
        for i, col_name in enumerate(df.columns):
            worksheet.write(0, i + 1, col_name)  # Write the header in B, C, D, ...
            for j, value in enumerate(df[col_name]):
                worksheet.write(j + 1, i + 1, value)  # Write the data in rows starting from B2

        # Apply the undrafted format to header row as well
        worksheet.set_row(0, None, undrafted_format)

        # Add a button in column A for each player
        for row_num in range(1, len(df) + 1):
            worksheet.insert_button(f'A{row_num + 1}', {
                'macro': 'ToggleDraftedStatus',
                'caption': 'Draft',
                'width': 50,
                'height': 20,
                'x_offset': 2,
                'y_offset': 2
            })

            # Apply the base format to each row initially
            worksheet.set_row(row_num, cell_format=base_format)

            # Apply a thick border to help estimate number of rounds
            if row_num % 12 == 0:
                worksheet.set_row(row_num, cell_format=round_end_format)

            # Apply conditional formatting based on the value in column A of the current row
            worksheet.conditional_format(f'B{row_num + 1}:Z{row_num + 1}', {
                'type': 'formula',
                'criteria': f'=LEN($A{row_num + 1})>0',
                'format': drafted_format
            })
            worksheet.conditional_format(f'B{row_num + 1}:Z{row_num + 1}', {
                'type': 'formula',
                'criteria': f'=LEN($A{row_num + 1})=0',
                'format': undrafted_format
            })

        # Apply position-specific colors to the "Position" column if it exists
        if "Pos." in df.columns:
            position_col_index = df.columns.get_loc("Pos.") + 1  # Adjust for xlsxwriter (1-based index)

            worksheet.conditional_format(1, position_col_index, len(df), position_col_index, {
                'type': 'cell',
                'criteria': 'equal to',
                'value': '"QB"',
                'format': qb_format
            })
            worksheet.conditional_format(1, position_col_index, len(df), position_col_index, {
                'type': 'cell',
                'criteria': 'equal to',
                'value': '"RB"',
                'format': rb_format
            })
            worksheet.conditional_format(1, position_col_index, len(df), position_col_index, {
                'type': 'cell',
                'criteria': 'equal to',
                'value': '"WR"',
                'format': wr_format
            })
            worksheet.conditional_format(1, position_col_index, len(df), position_col_index, {
                'type': 'cell',
                'criteria': 'equal to',
                'value': '"TE"',
                'format': te_format
            })

        # Set the width of the "Name" column to fit the longest name
        if "Name" in df.columns:
            max_name_length = df["Name"].str.len().max()
            worksheet.set_column(1, 1, max_name_length + 2)  # Column 'B' (index 1)

        # Apply conditional formatting to the "Team" column based on team rank (if they exist)
        if len(team_rankings) > 0:
            min_rank = min(team_rankings.values())
            max_rank = max(team_rankings.values())

            # Get 32 colors from GOOD_COLOR, NEUTRAL_COLOR, BAD_COLOR range
            good_rgb = (int(GOOD_COLOR[1:3], 16), int(GOOD_COLOR[3:5], 16), int(GOOD_COLOR[5:7], 16))
            neutral_rgb = (int(NEUTRAL_COLOR[1:3], 16), int(NEUTRAL_COLOR[3:5], 16), int(NEUTRAL_COLOR[5:7], 16))
            bad_rgb = (int(BAD_COLOR[1:3], 16), int(BAD_COLOR[3:5], 16), int(BAD_COLOR[5:7], 16))
            half_of_rankings = len(team_rankings) / 2
            step_btwn_good_and_neutral = ((neutral_rgb[0] - good_rgb[0]) / half_of_rankings,
                                          (neutral_rgb[1] - good_rgb[1]) / half_of_rankings,
                                          (neutral_rgb[2] - good_rgb[2]) / half_of_rankings,)
            step_btwn_neutral_and_bad = ((neutral_rgb[0] - bad_rgb[0]) / half_of_rankings,
                                          (neutral_rgb[1] - bad_rgb[1]) / half_of_rankings,
                                          (neutral_rgb[2] - bad_rgb[2]) / half_of_rankings,)
            color_range = []
            for i in range(len(team_rankings)):
                if i < half_of_rankings:
                    rgb = (round(good_rgb[0] + (i * step_btwn_good_and_neutral[0])),
                           round(good_rgb[1] + (i * step_btwn_good_and_neutral[1])),
                           round(good_rgb[2] + (i * step_btwn_good_and_neutral[2])))
                    color_range.append(f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}")
                elif i == half_of_rankings:
                    color_range.append(NEUTRAL_COLOR)
                else:
                    i -= round(half_of_rankings)
                    rgb = (round(neutral_rgb[0] - (i * step_btwn_neutral_and_bad[0])),
                           round(neutral_rgb[1] - (i * step_btwn_neutral_and_bad[1])),
                           round(neutral_rgb[2] - (i * step_btwn_neutral_and_bad[2])))
                    color_range.append(f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}")

            if "Team" in df.columns:
                team_col_index = df.columns.get_loc("Team") + 1  # Adjust for xlsxwriter (1-based index)
                for row_num in range(1, len(df) + 1):
                    team_name = df.at[row_num - 1, "Team"]  # Get the team name for this row
                    if team_name in team_rankings:
                        # Determine color based on rank
                        bg_color = color_range[team_rankings[team_name] - 1]
                        team_format = workbook.add_format({
                            'font_size': 14,
                            'bg_color': bg_color,
                            'bottom': 0 if row_num % 12 else 2
                        })
                        worksheet.write(row_num, team_col_index, team_name, team_format)

        # Apply heatmap to the columns except "Name", "Team", and "Position"
        for i, column_name in enumerate(df.columns):
            if column_name not in ["Name", "Team", "Position"]:
                col_index = i + 1  # Adjusted index for xlsxwriter
                if column_name == "Range":
                    worksheet.conditional_format(1, col_index, len(df), col_index, {
                        'type': '3_color_scale',
                        'min_type': 'percentile',
                        'min_value': 0,
                        'min_color': GOOD_COLOR,
                        'mid_type': 'percentile',
                        'mid_value': 50,
                        'mid_color': NEUTRAL_COLOR,
                        'max_type': 'percentile',
                        'max_value': 100,
                        'max_color': BAD_COLOR
                    })
                else:
                    worksheet.conditional_format(1, col_index, len(df), col_index, {
                        'type': '3_color_scale',
                        'min_type': 'min',
                        'min_color': BAD_COLOR,
                        'mid_type': 'num',
                        'mid_value': 0,
                        'mid_color': NEUTRAL_COLOR,
                        'max_type': 'max',
                        'max_color': GOOD_COLOR
                    })

    # Include the VBA macro script in your workbook
    workbook.add_vba_project('./vbaProject.bin')  # Assumes you have a vbaProject.bin file with the macro

    # Close the workbook (saves it)
    workbook.close()

    print("Done")


