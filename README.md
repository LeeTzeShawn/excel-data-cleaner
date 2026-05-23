# Excel Data Cleaner

Clean Excel files containing player data and win/loss records.

## How to Use

1. Click the green **"Code"** button above
2. Select **"Codespaces"** tab
3. Click **"Create codespace on main"**
4. Wait 1-2 minutes for setup
5. Upload your Excel file (drag and drop into the file explorer)
6. Open terminal (Ctrl+\` or Cmd+\`)
7. Run: `python clean_script.py your_file.xlsx`
8. Download the cleaned file: `your_file_cleaned.xlsx`

## Input Format Expected

Your Excel should have:
- **Table 1**: Player info (rows 0-4, columns A-E)
- **Table 2**: Win/loss records (rows 10-15, columns G-I)

## Output

Creates a cleaned Excel file with:
- Player info
- Win/loss records
- Win percentage
- Total matches
- Win/loss ratio
- Rank by winsR
