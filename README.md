# FIFA World Cup Sweepstake Website

This website reads data directly from the Excel file:

`FIFA World Cup Leaderboard - New.xlsx`

No API is used.

## GitHub Pages setup

1. Upload all files in this folder to your GitHub repository.
2. Keep the Excel file in the same folder as `index.html`.
3. In GitHub, go to **Settings > Pages**.
4. Select the branch and root folder, then save.
5. Open your GitHub Pages URL.

## Updating scores

1. Open `FIFA World Cup Leaderboard - New.xlsx`.
2. Update the `Participants` sheet and/or `standings` sheet.
3. Save the file with the same name.
4. Upload/replace the Excel file in GitHub.
5. Refresh the website.

## Expected Excel sheets

### Participants
Required columns:
- Participant
- Team 1
- Team 2
- Team 3
- Total Points

### standings
Required columns:
- Group
- Position
- Team
- Code
- Played
- Won
- Draw
- Lost
- Points
- GoalsFor
- GoalsAgainst
- GoalDifference
