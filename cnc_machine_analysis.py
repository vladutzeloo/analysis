#!/usr/bin/env python3
"""
CNC Machine Running Time Analysis
Extracts production data from HTML files and calculates actual running hours per machine
Excludes sample parts (cycle time = 480s)
"""

import os
import re
import sys
from datetime import datetime
from pathlib import Path
from html.parser import HTMLParser
from collections import defaultdict
import json


class ProductionHTMLParser(HTMLParser):
    """Custom HTML parser to extract production data from dashboard tables"""

    def __init__(self):
        super().__init__()
        self.in_tbody = False
        self.in_tr = False
        self.in_td = False
        self.current_td_class = None
        self.current_row = []
        self.current_cell = []
        self.td_count = 0
        self.data = []

    def handle_starttag(self, tag, attrs):
        attrs_dict = dict(attrs)

        if tag == 'tbody':
            self.in_tbody = True
        elif tag == 'tr' and self.in_tbody:
            self.in_tr = True
            self.current_row = []
            self.td_count = 0
        elif tag == 'td' and self.in_tr:
            self.in_td = True
            self.current_cell = []
            self.current_td_class = attrs_dict.get('class', '')

    def handle_endtag(self, tag):
        if tag == 'tbody':
            self.in_tbody = False
        elif tag == 'tr' and self.in_tr:
            self.in_tr = False
            # Only add rows that have the expected number of columns (11 columns in the table)
            if len(self.current_row) >= 10:
                self.data.append(self.current_row)
        elif tag == 'td' and self.in_td:
            self.in_td = False
            cell_text = ''.join(self.current_cell).strip()
            self.current_row.append({
                'text': cell_text,
                'class': self.current_td_class
            })
            self.td_count += 1

    def handle_data(self, data):
        if self.in_td:
            self.current_cell.append(data)


def parse_cycle_time(cycle_time_str):
    """
    Parse cycle time string and convert to seconds
    Examples: "1.5s" -> 1.5, "480.0s" -> 480.0, "—" -> None
    """
    if not cycle_time_str or cycle_time_str.strip() in ['—', '-', 'N/A', '']:
        return None

    match = re.search(r'([\d.]+)\s*s', cycle_time_str)
    if match:
        return float(match.group(1))
    return None


def extract_production_data(html_file):
    """
    Extract production data from a single HTML file
    Returns list of records with machine, item, parts, cycle time, etc.
    """
    try:
        with open(html_file, 'r', encoding='utf-8') as f:
            html_content = f.read()

        # Extract date from filename if possible
        date_match = re.search(r'(\d{4})(\d{2})(\d{2})', str(html_file))
        if date_match:
            file_date = f"{date_match.group(1)}-{date_match.group(2)}-{date_match.group(3)}"
        else:
            file_date = "Unknown"

        parser = ProductionHTMLParser()
        parser.feed(html_content)

        records = []
        for row in parser.data:
            if len(row) < 10:
                continue

            # Extract data from columns
            # Column structure: Machine, Operation, Item, Order, OK Parts, NOK Parts, Quality, Cycle Time, Setup, OEE, Operator
            machine = row[0]['text']
            operation = row[1]['text'] if len(row) > 1 else ''
            item = row[2]['text'] if len(row) > 2 else ''
            order = row[3]['text'] if len(row) > 3 else ''

            # Parse OK parts (column 4)
            try:
                ok_parts = int(row[4]['text'].replace(',', ''))
            except (ValueError, IndexError):
                ok_parts = 0

            # Parse NOK parts (column 5)
            try:
                nok_parts = int(row[5]['text'].replace(',', ''))
            except (ValueError, IndexError):
                nok_parts = 0

            # Parse cycle time (column 7)
            cycle_time_str = row[7]['text'] if len(row) > 7 else ''
            cycle_time = parse_cycle_time(cycle_time_str)

            # Extract operator and shift
            operator = row[10]['text'] if len(row) > 10 else ''
            shift_match = re.search(r'\(S(\d)\)', operator)
            shift = f"S{shift_match.group(1)}" if shift_match else "Unknown"

            if machine and cycle_time is not None:
                records.append({
                    'date': file_date,
                    'machine': machine,
                    'operation': operation,
                    'item': item,
                    'order': order,
                    'ok_parts': ok_parts,
                    'nok_parts': nok_parts,
                    'cycle_time': cycle_time,
                    'operator': operator,
                    'shift': shift,
                    'is_sample': cycle_time == 480.0  # Flag sample parts
                })

        return records

    except Exception as e:
        print(f"Error parsing {html_file}: {e}")
        return []


def calculate_running_hours(records, exclude_samples=True):
    """
    Calculate actual running hours per machine from production records

    Args:
        records: List of production records
        exclude_samples: If True, exclude records where cycle_time = 480s

    Returns:
        Dictionary with machine statistics
    """
    machine_stats = defaultdict(lambda: {
        'total_parts': 0,
        'total_parts_with_samples': 0,
        'sample_parts': 0,
        'total_seconds': 0.0,
        'total_seconds_with_samples': 0.0,
        'sample_seconds': 0.0,
        'records': [],
        'dates': set(),
        'items': set(),
        'shifts': set()
    })

    for record in records:
        machine = record['machine']
        ok_parts = record['ok_parts']
        cycle_time = record['cycle_time']
        is_sample = record['is_sample']

        # Calculate running time in seconds
        running_seconds = ok_parts * cycle_time

        # Always track total with samples
        machine_stats[machine]['total_parts_with_samples'] += ok_parts
        machine_stats[machine]['total_seconds_with_samples'] += running_seconds
        machine_stats[machine]['dates'].add(record['date'])
        machine_stats[machine]['items'].add(record['item'])
        machine_stats[machine]['shifts'].add(record['shift'])

        # Track samples separately
        if is_sample:
            machine_stats[machine]['sample_parts'] += ok_parts
            machine_stats[machine]['sample_seconds'] += running_seconds
        else:
            # Only count production parts if not excluding samples
            machine_stats[machine]['total_parts'] += ok_parts
            machine_stats[machine]['total_seconds'] += running_seconds

        machine_stats[machine]['records'].append(record)

    # Convert seconds to hours
    for machine, stats in machine_stats.items():
        stats['total_hours'] = stats['total_seconds'] / 3600
        stats['total_hours_with_samples'] = stats['total_seconds_with_samples'] / 3600
        stats['sample_hours'] = stats['sample_seconds'] / 3600
        stats['dates'] = sorted(list(stats['dates']))
        stats['items'] = sorted(list(stats['items']))
        stats['shifts'] = sorted(list(stats['shifts']))

    return dict(machine_stats)


def generate_report(machine_stats, output_file=None):
    """
    Generate a comprehensive report of machine running hours
    """
    report_lines = []

    # Header
    report_lines.append("=" * 100)
    report_lines.append("CNC MACHINE ACTUAL RUNNING TIME ANALYSIS REPORT")
    report_lines.append("=" * 100)
    report_lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report_lines.append("")

    # Summary statistics
    total_machines = len(machine_stats)
    total_production_hours = sum(stats['total_hours'] for stats in machine_stats.values())
    total_sample_hours = sum(stats['sample_hours'] for stats in machine_stats.values())
    total_all_hours = sum(stats['total_hours_with_samples'] for stats in machine_stats.values())

    report_lines.append("SUMMARY")
    report_lines.append("-" * 100)
    report_lines.append(f"Total Machines: {total_machines}")
    report_lines.append(f"Total Production Hours (excluding samples): {total_production_hours:.2f} hours")
    report_lines.append(f"Total Sample Hours (cycle time = 480s): {total_sample_hours:.2f} hours")
    report_lines.append(f"Total All Hours (including samples): {total_all_hours:.2f} hours")
    report_lines.append("")

    # Sort machines by total production hours (descending)
    sorted_machines = sorted(
        machine_stats.items(),
        key=lambda x: x[1]['total_hours'],
        reverse=True
    )

    # Detailed machine breakdown
    report_lines.append("DETAILED MACHINE BREAKDOWN")
    report_lines.append("-" * 100)
    report_lines.append("")

    for machine, stats in sorted_machines:
        report_lines.append(f"Machine: {machine}")
        report_lines.append(f"  Production Hours (excl. samples): {stats['total_hours']:.2f} hours")
        report_lines.append(f"  Sample Hours (480s cycle time):   {stats['sample_hours']:.2f} hours")
        report_lines.append(f"  Total Hours (incl. samples):      {stats['total_hours_with_samples']:.2f} hours")
        report_lines.append(f"  Production Parts:                 {stats['total_parts']:,}")
        report_lines.append(f"  Sample Parts:                     {stats['sample_parts']:,}")
        report_lines.append(f"  Total Parts:                      {stats['total_parts_with_samples']:,}")
        report_lines.append(f"  Dates Active:                     {', '.join(stats['dates'])}")
        report_lines.append(f"  Shifts:                           {', '.join(stats['shifts'])}")
        report_lines.append(f"  Items Produced:                   {len(stats['items'])}")

        # Show sample items if any
        sample_records = [r for r in stats['records'] if r['is_sample']]
        if sample_records:
            report_lines.append(f"  Sample Production Records:")
            for rec in sample_records:
                report_lines.append(f"    - {rec['date']} | {rec['shift']} | {rec['item']} | {rec['ok_parts']} parts")

        report_lines.append("")

    # Item-level breakdown
    report_lines.append("=" * 100)
    report_lines.append("ITEM-LEVEL PRODUCTION SUMMARY")
    report_lines.append("-" * 100)
    report_lines.append("")

    # Collect all unique items across all machines
    item_stats = defaultdict(lambda: {
        'machines': set(),
        'total_parts': 0,
        'total_hours': 0.0,
        'is_sample': False,
        'cycle_times': set()
    })

    for machine, stats in machine_stats.items():
        for record in stats['records']:
            item = record['item']
            item_stats[item]['machines'].add(machine)
            item_stats[item]['cycle_times'].add(record['cycle_time'])
            if not record['is_sample']:
                item_stats[item]['total_parts'] += record['ok_parts']
                item_stats[item]['total_hours'] += (record['ok_parts'] * record['cycle_time']) / 3600
            if record['is_sample']:
                item_stats[item]['is_sample'] = True

    sorted_items = sorted(item_stats.items(), key=lambda x: x[1]['total_hours'], reverse=True)

    for item, stats in sorted_items:
        cycle_times_str = ', '.join(f"{ct}s" for ct in sorted(stats['cycle_times']))
        sample_flag = " [SAMPLE]" if stats['is_sample'] else ""
        report_lines.append(f"Item: {item}{sample_flag}")
        report_lines.append(f"  Production Hours: {stats['total_hours']:.2f} hours")
        report_lines.append(f"  Total Parts:      {stats['total_parts']:,}")
        report_lines.append(f"  Machines:         {', '.join(sorted(stats['machines']))}")
        report_lines.append(f"  Cycle Times:      {cycle_times_str}")
        report_lines.append("")

    # Footer
    report_lines.append("=" * 100)
    report_lines.append("NOTE: Sample parts (cycle time = 480s) are tracked separately and excluded from production hours")
    report_lines.append("=" * 100)

    report_text = '\n'.join(report_lines)

    # Print to console
    print(report_text)

    # Save to file if specified
    if output_file:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(report_text)
        print(f"\n✅ Report saved to: {output_file}")

    return report_text


def main():
    """Main function to analyze CNC machine running times"""
    print("🏭 CNC Machine Running Time Analysis")
    print("=" * 70)
    print()

    # Get folder path from command line or prompt
    if len(sys.argv) > 1:
        folder_path = sys.argv[1]
    else:
        folder_path = input("Enter the folder path containing HTML files: ").strip()

    if not os.path.exists(folder_path):
        print(f"❌ Error: Folder not found: {folder_path}")
        return

    # Find all HTML files
    html_files = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.html'):
                html_files.append(os.path.join(root, file))

    if not html_files:
        print(f"❌ No HTML files found in: {folder_path}")
        return

    print(f"📁 Found {len(html_files)} HTML file(s)")
    print()

    # Parse all HTML files
    all_records = []
    for html_file in html_files:
        print(f"📄 Processing: {os.path.basename(html_file)}")
        records = extract_production_data(html_file)
        all_records.extend(records)
        print(f"   ✓ Extracted {len(records)} records")

    print()
    print(f"✅ Total records extracted: {len(all_records)}")
    print()

    # Calculate running hours
    machine_stats = calculate_running_hours(all_records, exclude_samples=True)

    # Generate report
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"cnc_running_hours_report_{timestamp}.txt"
    generate_report(machine_stats, output_file)

    # Also save JSON data for further analysis
    json_file = f"cnc_running_hours_data_{timestamp}.json"

    # Convert sets to lists for JSON serialization
    json_data = {}
    for machine, stats in machine_stats.items():
        json_data[machine] = {
            'total_hours': stats['total_hours'],
            'total_hours_with_samples': stats['total_hours_with_samples'],
            'sample_hours': stats['sample_hours'],
            'total_parts': stats['total_parts'],
            'sample_parts': stats['sample_parts'],
            'dates': stats['dates'],
            'items': stats['items'],
            'shifts': stats['shifts']
        }

    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(json_data, f, indent=2)

    print(f"✅ JSON data saved to: {json_file}")
    print()
    print("=" * 70)
    print("Analysis complete!")


if __name__ == "__main__":
    main()
