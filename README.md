# Fantasy Football Draft Sheet

This script will take fantasy football projections and create
a custom draft sheet based on those projections.

## Usage instructions

1. Put CSV files with stat projections in the `fp_data` folder.
   These CSV files should follow the naming convention of ending
   with an underscore followed by the position, for example `Data_QB.csv`.
   The data should be formatted similar to how FantasyPros data looks,
   except with slightly different headers that can be seen at the top
   of the code (`main.py`). It is recommended to provide low, average, and 
   high projections, however only the average projections are absolutely necessary.
2. Put a CSV file with the rankings from your drafting platform in the 
   `draft_platform_rankings` folder. The CSV file should have 3 columns: `Name`,
   `Position`, and the name of the platform which will show up on the final draft sheet.
   The column containing the name of the platform should hold that player's overall
   rank on that platform. **You can skip this step, but** the related column
   will be missing on your draft sheet, so it isn't recommended.
3. (Optional) Add a file `teams.csv` in the `notes` folder. It should have a sorted
   list of teams in the first column. This sorting determines what rank you think each team
   is, and the team will be colored appropriately on the draft sheet to indicate their strength. 
4. Run `main.py`. Don't forget to install the requirements 
   (`pip install -r requirements.txt`)
5. Open `DraftSheet.xlsm` (enable macros) and you're ready to draft!
   It's normal for the spreadsheet to take a long time to load when
   first opening it, but once it's loaded it'll be quick.

## Draft time: using the sheet effectively

Let's start with the positional ranking sheets (QB, TE, RB, WR).
The `low`, `avg`, and `high` columns show low, average, and high projections
for per-week scoring compared to the baseline player. The baseline player
is the projected player who would be drafted after the 9th round. 

For example, let's consider the case where Bijan Robinson has a 
low, average, and high of 6.6, 9.8, and 13.2, respectively.
This means he is projected to score at least 6.6 more points per week than
the baseline player, at most 13.2 points per week more than the baseline
player, and on average 9.8 points per week more than the baseline player.
Remember, these are based on projections! If you want to know who the
baseline player is, look for the player with an `avg` of `0`.

The range for each player is there to help decipher the volatility of a player.
It is calculated by subtracting the player's lowest projection from their highest
projection of points scored per week. Players with a lower range have more
consistent projections, while players with a higher range have more volatile
projections. 

The last column, which should have the same name as your drafting plaform,
shows the difference between your projected rank and the platform's projected rank.
This can help you know if you can wait another round before a player shows up on 
your opponents' screens (assuming they don't have their own top-tier draft sheet)!
If the difference is -9999, this is a special case where that player is not ranked on
the drafting platform (or the data could be missing from the CSV file).

The `Overall` sheet combines all the position info to give an overall draft
ranking. The low, average, and high projections are still relative to the baseline
player of the same position. In other words, the draft sheet is ordered by value over
replacement, or by value considering positional scarcity. 

When you draft a player, hit the draft button, and it will update on all sheets.
If you mistakenly click a player, clicking the button again will return them to
the undrafted appearance. 

## FAQ

**Do I need to enable macros?**

No, but the buttons to mark a player as drafted will not work.
There are 2 workarounds for this:
1. Write any value in the A column of a player that has been drafted.
   Underneath the button, the cell still exists, and you can write in it.
   Clearing that cell will mark the player as undrafted (default state).
2. Load the macro manually. Copy/paste the script saved in `ToggleDraftStatus.vba`.

