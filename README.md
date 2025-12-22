# Mizrahi Fund Automation

Automated fund report processing and validation system for Mizrahi Tefahot trustee services. This tool fetches fund data from Maya (TASE) via Apify actors and performs comprehensive compliance checks.

## Features

- **Automated Data Fetching**: Retrieves fund lists and manager reports from Maya/TASE using Apify actors
- **Fund Completeness Check**: Cross-references funds between Magna list and manager reports
- **Unusual Asset Detection**: Flags holdings with unusual asset types (types 16, 21, 22, 23, 24, 52, 53, 57, 58, 99, 101, 112, 201, 207, 209)
- **Change Tracking**: Identifies new assets and quantity changes from previous month
- **Regulatory Compliance**: Validates Clause 214 (variable management fees) and Clause 328 (borrowed quantities)
- **Required Combinations**: Verifies required asset type pairings exist
- **Price Reasonableness**: Checks price ratios with 7.5% threshold

## Supported Fund Managers

- מגדל (Migdal)
- איילון (Ayalon)
- קסם (Kesem)
- סיגמא (Sigma)
- פורסט (Forest)
- הראל (Harel)
- אנליסט (Analyst)
- מיטב (Meitav)
- איביאי (IBI)
- אלטשולר-שחם (Altshuler Shaham)

## Installation

```bash
pip install pandas openpyxl requests
```

## Configuration

Before running, set your Apify API token in the script:

```python
APIFY_TOKEN = "YOUR_APIFY_TOKEN_HERE"
```

## Usage

### Full Automation (Complete Pipeline)

```bash
python fund_automation_complete.py --fund-name "סיגמא"
python fund_automation_complete.py --fund-name "סיגמא" --output-dir ./reports
python fund_automation_complete.py --fund-name "סיגמא" --keep-temp  # Keep temp files
```

### Test Script (Apify Integration Test)

```bash
python test_fund_automation.py
```

This will:
1. Fetch the master funds list from Apify
2. Fetch fund reports (current and previous month CSVs)
3. Save files to `test_output/` directory

## Output

The main script generates an Excel report with the following sheets:

| Sheet | Description |
|-------|-------------|
| סיכום | Summary statistics |
| סטטוס בדיקות | Check status overview |
| קרנות חסרות | Missing funds (if any) |
| נכסים חריגים | Unusual assets |
| נכסים חדשים | New assets |
| שינויים בכמות | Quantity changes |
| סעיף 328 | Clause 328 issues |
| שילובים נדרשים | Required combinations (including Clause 214) |
| סבירות מחירים | Price reasonableness issues |

## Checks Performed

1. **Fund Completeness** - Cross-reference Magna funds vs manager report
2. **Unusual Asset Types** - Flag holdings with non-standard asset types
3. **New Assets** - Identify assets added since previous month
4. **Quantity Changes** - Track changes in unusual asset quantities
5. **Clause 328** - Verify borrowed quantity consistency
6. **Required Combinations** - Validate asset type pairings (111, 212, 213, 208, 210)
7. **Price Reasonableness** - Check price ratios (300/301, 314/313, 316/315)

## License

Proprietary - Mizrahi Tefahot
