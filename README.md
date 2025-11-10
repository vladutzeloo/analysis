# CNC Machine Running Time Analysis

This Python script analyzes production HTML files and calculates actual running hours for each CNC machine.

## Features

- ✅ Parses multiple HTML production dashboard files from a folder
- ✅ Extracts machine data, cycle times, and parts produced
- ✅ Calculates actual running hours per machine
- ✅ **Automatically excludes sample parts** (cycle time = 480s)
- ✅ Generates comprehensive text report and JSON data file
- ✅ Provides machine-level and item-level breakdowns
- ✅ Tracks production across multiple dates and shifts

## Key Logic

**Running Time Calculation:**
```
Actual Running Hours = (OK Parts × Cycle Time) / 3600
```

**Sample Parts Handling:**
- When `cycle time = 480s`, the machine was producing **sample parts**
- These are tracked separately and **excluded** from production hours
- The report shows both production hours (excluding samples) and total hours (including samples)

## Usage

### Run from command line:
```bash
python3 cnc_machine_analysis.py /path/to/html/files/folder
```

### Or run interactively:
```bash
python3 cnc_machine_analysis.py
```
Then enter the folder path when prompted.

## Output Files

The script generates three files with timestamps:

1. **HTML Dashboard** (`cnc_running_hours_dashboard_YYYYMMDD_HHMMSS.html`) ⭐ **Main Output**
   - Beautiful interactive dashboard with modern UI
   - Stacked bar chart showing production vs sample hours for all machines
   - Machine selector dropdown for detailed breakdowns
   - Filterable item-level production table
   - Interactive tooltips and real-time filtering
   - Summary cards with key metrics
   - Responsive design that works on any screen size

2. **Text Report** (`cnc_running_hours_report_YYYYMMDD_HHMMSS.txt`)
   - Summary statistics
   - Detailed machine breakdown with production/sample hours
   - Item-level production summary
   - Sample production records flagged

3. **JSON Data** (`cnc_running_hours_data_YYYYMMDD_HHMMSS.json`)
   - Machine-level statistics in structured format
   - Easy to import into Excel, Power BI, or other analysis tools

## Report Sections

### HTML Dashboard Sections:

#### 1. Summary Cards
- Production Hours (excluding samples) - Green card
- Sample Hours (480s cycle only) - Orange card
- Total Hours (all production) - Blue card
- Total Machines & Parts - Purple card

#### 2. Machine Running Hours Chart
- Interactive stacked bar chart
- Production hours (green) vs Sample hours (orange)
- Hover tooltips with detailed information
- Visual comparison of all machines

#### 3. Machine Details Breakdown
- **Dropdown selector** to choose a machine
- Detailed statistics grid:
  - Production hours, Sample hours, Total hours
  - Production parts, Sample parts
  - Items produced count
- Active dates, shifts, and items lists
- **Production Records Table** with:
  - Date, Shift, Item, Operation
  - Parts produced, Cycle time, Hours calculated
  - Type badge (SAMPLE/PRODUCTION)
  - Operator information

#### 4. Item-Level Production Summary
- **Filterable table** (search by item name)
- Production hours per item
- Total parts produced
- Number of machines used
- Cycle times
- Type badge (SAMPLE/PRODUCTION)

### Text Report Sections:

#### 1. Summary
- Total machines analyzed
- Total production hours (excluding samples)
- Total sample hours (cycle time = 480s)
- Total all hours (including samples)

#### 2. Machine Breakdown
For each machine:
- Production hours (excluding sample parts)
- Sample hours (480s cycle time only)
- Total hours (all production)
- Part counts (production vs sample)
- Active dates and shifts
- Items produced
- **List of sample production records** (if any)

#### 3. Item-Level Summary
For each item/part number:
- Total production hours
- Total parts produced
- Machines used
- Cycle times
- Sample flag if applicable

## Example Output

```
SUMMARY
----------------------------------------------------------------------------------------------------
Total Machines: 11
Total Production Hours (excluding samples): 1.93 hours
Total Sample Hours (cycle time = 480s): 0.53 hours
Total All Hours (including samples): 2.46 hours

Machine: 211 - SW
  Production Hours (excl. samples): 0.12 hours
  Sample Hours (480s cycle time):   0.53 hours  ← Sample parts detected
  Total Hours (incl. samples):      0.65 hours
  Production Parts:                 103
  Sample Parts:                     4            ← 4 sample parts
  Sample Production Records:
    - 2025-11-07 | S1 | Item-0 | 4 parts       ← Details of sample production
```

## Requirements

- Python 3.6 or higher
- No external dependencies (uses only standard library)

## Supported HTML Format

The script expects HTML files with production dashboard tables containing these columns:
- Machine
- Operation
- Item
- Order
- OK Parts
- NOK Parts
- Quality %
- **Cycle Time** (format: "X.Xs" where 480.0s indicates samples)
- Setup
- OEE %
- Operator (with shift information)

## Notes

- Files should follow naming pattern: `olstral_production_dashboard*.html`
- Dates are extracted from filenames (format: YYYYMMDD)
- All cycle times are in seconds (converted from "Xs" format)
- Sample parts (480s cycle time) are completely separated in calculations
- JSON output can be used for further analysis in Excel, Python pandas, etc.

## Troubleshooting

**No records found?**
- Check that HTML files contain production tables
- Verify the table structure matches expected format

**Wrong cycle times?**
- Ensure cycle times in HTML are in format "X.Xs" (e.g., "1.5s", "480.0s")

**Sample parts not excluded?**
- Verify that sample cycle times are exactly "480.0s" in the HTML
