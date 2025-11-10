#!/usr/bin/env python3
"""
CNC Machine Running Time Analysis
Extracts production data from HTML files and calculates actual running hours per machine
Note: Cycle times in HTML are in MINUTES (e.g., "1.5s" = 1.5 minutes)
Excludes sample parts (cycle time = 480 minutes = 8 hours)
"""

import os
import re
import sys
from datetime import datetime
from pathlib import Path
from html.parser import HTMLParser
from collections import defaultdict
import json

# Try to import tkinter for GUI dialogs, but make it optional
try:
    import tkinter as tk
    from tkinter import filedialog
    TKINTER_AVAILABLE = True
except ImportError:
    TKINTER_AVAILABLE = False


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
    Note: The values in HTML are in MINUTES, not seconds!
    Examples: "1.5s" -> 1.5 minutes -> 90 seconds
              "480.0s" -> 480 minutes -> 28,800 seconds (8 hours)
    """
    if not cycle_time_str or cycle_time_str.strip() in ['—', '-', 'N/A', '']:
        return None

    match = re.search(r'([\d.]+)\s*s', cycle_time_str)
    if match:
        # The value is in MINUTES, convert to seconds
        minutes = float(match.group(1))
        return minutes * 60  # Convert minutes to seconds
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
                    'is_sample': cycle_time == 28800.0  # Flag sample parts (480 minutes = 28,800 seconds)
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
        exclude_samples: If True, exclude records where cycle_time = 480 minutes (28,800 seconds)

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
        'has_samples': False,
        'records': [],
        'dates': set(),
        'items': set(),
        'shifts': set(),
        'shift_counts': defaultdict(int)  # Track how many times each shift appears
    })

    for record in records:
        machine = record['machine']
        ok_parts = record['ok_parts']
        cycle_time = record['cycle_time']
        is_sample = record['is_sample']
        shift = record['shift']

        # Track shift occurrences for downtime calculation
        machine_stats[machine]['shift_counts'][shift] += 1

        # Calculate running time in seconds
        # For sample parts (480 min), just count 480 minutes ONCE, not per part
        if is_sample:
            running_seconds = 0  # Don't add to running time
            if not machine_stats[machine]['has_samples']:
                # First time seeing samples for this machine, add 480 min = 8 hours
                machine_stats[machine]['sample_seconds'] = 28800.0  # 480 min * 60 sec = 28,800 sec
                machine_stats[machine]['has_samples'] = True
            machine_stats[machine]['sample_parts'] += ok_parts
        else:
            running_seconds = ok_parts * cycle_time
            machine_stats[machine]['total_parts'] += ok_parts
            machine_stats[machine]['total_seconds'] += running_seconds

        # Always track total with samples
        machine_stats[machine]['total_parts_with_samples'] += ok_parts
        machine_stats[machine]['dates'].add(record['date'])
        machine_stats[machine]['items'].add(record['item'])
        machine_stats[machine]['shifts'].add(record['shift'])
        machine_stats[machine]['records'].append(record)

    # Convert seconds to hours and calculate downtime
    for machine, stats in machine_stats.items():
        stats['total_hours'] = stats['total_seconds'] / 3600
        stats['sample_hours'] = stats['sample_seconds'] / 3600
        stats['total_hours_with_samples'] = (stats['total_seconds'] + stats['sample_seconds']) / 3600

        # Calculate downtime: Available capacity - actual running time
        # Each shift = 8 hours available
        total_shift_occurrences = sum(stats['shift_counts'].values())
        available_hours = total_shift_occurrences * 8  # 8 hours per shift
        stats['available_capacity'] = available_hours
        stats['downtime_hours'] = max(0, available_hours - stats['total_hours_with_samples'])
        stats['availability_percent'] = (stats['total_hours_with_samples'] / available_hours * 100) if available_hours > 0 else 0
        stats['downtime_percent'] = 100 - stats['availability_percent']

        stats['dates'] = sorted(list(stats['dates']))
        stats['items'] = sorted(list(stats['items']))
        stats['shifts'] = sorted(list(stats['shifts']))

    return dict(machine_stats)


def generate_html_report(machine_stats, all_records, output_file, available_months):
    """
    Generate a beautiful interactive HTML report with charts and filtering
    """
    # Calculate summary statistics
    total_machines = len(machine_stats)
    total_production_hours = sum(stats['total_hours'] for stats in machine_stats.values())
    total_sample_hours = sum(stats['sample_hours'] for stats in machine_stats.values())
    total_all_hours = sum(stats['total_hours_with_samples'] for stats in machine_stats.values())
    total_production_parts = sum(stats['total_parts'] for stats in machine_stats.values())
    total_sample_parts = sum(stats['sample_parts'] for stats in machine_stats.values())

    # Downtime statistics
    total_downtime_hours = sum(stats['downtime_hours'] for stats in machine_stats.values())
    total_available_hours = sum(stats['available_capacity'] for stats in machine_stats.values())
    overall_availability = (total_all_hours / total_available_hours * 100) if total_available_hours > 0 else 0
    overall_downtime_percent = 100 - overall_availability

    # Get all unique dates
    all_dates = set()
    for stats in machine_stats.values():
        all_dates.update(stats['dates'])
    date_range = f"{min(all_dates)} to {max(all_dates)}" if all_dates else "N/A"

    # Sort machines by production hours
    sorted_machines = sorted(
        machine_stats.items(),
        key=lambda x: x[1]['total_hours'],
        reverse=True
    )

    # Prepare data for charts
    machine_names = [m[0] for m in sorted_machines]
    production_hours = [m[1]['total_hours'] for m in sorted_machines]
    sample_hours = [m[1]['sample_hours'] for m in sorted_machines]

    # Collect item stats
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

    # Generate HTML
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CNC Machine Running Time Analysis - {datetime.now().strftime('%Y-%m-%d')}</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/3.9.1/chart.min.js"></script>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}

        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #1a1a1a 0%, #2d1b1b 100%);
            min-height: 100vh;
            padding: 20px;
            color: #ffffff;
        }}

        .container {{
            max-width: 1600px;
            margin: 0 auto;
            background: linear-gradient(135deg, #2d1b1b 0%, #1a1a1a 100%);
            border-radius: 20px;
            box-shadow: 0 20px 40px rgba(220, 38, 38, 0.3);
            overflow: hidden;
            border: 1px solid rgba(220, 38, 38, 0.3);
        }}

        .header {{
            background: linear-gradient(135deg, #dc2626 0%, #991b1b 100%);
            color: white;
            padding: 30px;
            text-align: center;
            position: relative;
        }}

        .header h1 {{
            font-size: 2.5rem;
            margin: 0;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }}

        .header .subtitle {{
            font-size: 1.2rem;
            opacity: 0.9;
            margin-top: 10px;
            background: rgba(255, 255, 255, 0.1);
            padding: 8px 16px;
            border-radius: 6px;
            display: inline-block;
        }}

        .summary-cards {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            padding: 30px;
            background: rgba(45, 27, 27, 0.5);
        }}

        .summary-card {{
            background: linear-gradient(135deg, #2d1b1b 0%, #1a1a1a 100%);
            padding: 25px;
            border-radius: 15px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.3);
            text-align: center;
            position: relative;
            border: 1px solid rgba(220, 38, 38, 0.3);
            transition: transform 0.3s ease;
        }}

        .summary-card::before {{
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 4px;
            background: var(--card-color);
        }}

        .summary-card:hover {{
            transform: translateY(-5px);
        }}

        .summary-card h3 {{
            color: rgba(255, 255, 255, 0.7);
            font-size: 0.9rem;
            font-weight: 600;
            text-transform: uppercase;
            margin-bottom: 10px;
        }}

        .summary-card .value {{
            font-size: 2.5rem;
            font-weight: 700;
            color: var(--card-color);
            margin-bottom: 5px;
        }}

        .summary-card .label {{
            color: rgba(255, 255, 255, 0.5);
            font-size: 0.9rem;
        }}

        .summary-card.production {{ --card-color: #dc2626; }}
        .summary-card.sample {{ --card-color: #f59e0b; }}
        .summary-card.total {{ --card-color: #059669; }}
        .summary-card.parts {{ --card-color: #ea580c; }}

        .content-section {{
            padding: 30px;
        }}

        .section-title {{
            font-size: 1.8rem;
            margin-bottom: 20px;
            color: #dc2626;
            border-bottom: 2px solid #dc2626;
            padding-bottom: 10px;
        }}

        .chart-container {{
            background: linear-gradient(135deg, #2d1b1b 0%, #1a1a1a 100%);
            padding: 25px;
            border-radius: 15px;
            box-shadow: 0 4px 8px rgba(0,0,0,0.3);
            border: 1px solid rgba(220, 38, 38, 0.3);
            margin-bottom: 30px;
        }}

        .chart-container h3 {{
            color: #ffffff;
            font-size: 1.3rem;
            margin-bottom: 20px;
            text-align: center;
        }}

        .chart-wrapper {{
            position: relative;
            height: 400px;
            max-height: 500px;
            overflow-x: auto;
            overflow-y: hidden;
        }}

        canvas {{
            max-height: 400px;
        }}

        .machine-selector {{
            width: 100%;
            padding: 15px;
            font-size: 1rem;
            border-radius: 10px;
            border: 2px solid #dc2626;
            background: rgba(255, 255, 255, 0.1);
            color: white;
            margin-bottom: 20px;
            cursor: pointer;
        }}

        .machine-selector option {{
            background: #16213e;
            color: white;
        }}

        .machine-details {{
            background: rgba(255, 255, 255, 0.05);
            padding: 25px;
            border-radius: 15px;
            margin-bottom: 20px;
        }}

        .machine-details h3 {{
            color: #dc2626;
            margin-bottom: 15px;
            font-size: 1.4rem;
        }}

        .detail-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin-bottom: 20px;
        }}

        .detail-item {{
            background: rgba(14, 165, 233, 0.1);
            padding: 15px;
            border-radius: 10px;
            border-left: 4px solid #dc2626;
        }}

        .detail-item .label {{
            font-size: 0.85rem;
            opacity: 0.8;
            margin-bottom: 5px;
        }}

        .detail-item .value {{
            font-size: 1.5rem;
            font-weight: bold;
            color: #dc2626;
        }}

        .records-table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }}

        .records-table th {{
            background: rgba(14, 165, 233, 0.2);
            padding: 12px;
            text-align: left;
            font-weight: 600;
            border-bottom: 2px solid #dc2626;
        }}

        .records-table td {{
            padding: 12px;
            border-bottom: 1px solid rgba(255, 255, 255, 0.1);
        }}

        .records-table tr:hover {{
            background: rgba(14, 165, 233, 0.1);
        }}

        .sample-badge {{
            background: #f59e0b;
            color: white;
            padding: 4px 8px;
            border-radius: 5px;
            font-size: 0.8rem;
            font-weight: bold;
        }}

        .production-badge {{
            background: #059669;
            color: white;
            padding: 4px 8px;
            border-radius: 5px;
            font-size: 0.8rem;
            font-weight: bold;
        }}

        .filter-container {{
            display: flex;
            gap: 15px;
            margin-bottom: 20px;
            flex-wrap: wrap;
        }}

        .filter-input {{
            flex: 1;
            min-width: 200px;
            padding: 12px;
            border-radius: 10px;
            border: 2px solid #dc2626;
            background: rgba(255, 255, 255, 0.1);
            color: white;
            font-size: 1rem;
        }}

        .filter-input::placeholder {{
            color: rgba(255, 255, 255, 0.5);
        }}

        .alert-box {{
            background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%);
            color: white;
            padding: 20px;
            border-radius: 10px;
            margin-bottom: 20px;
            font-weight: 500;
        }}

        .footer {{
            background: rgba(15, 52, 96, 0.5);
            padding: 20px;
            text-align: center;
            font-size: 0.9rem;
            opacity: 0.8;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>⚙️ CNC Machine Running Time Analysis - Power BI Dashboard</h1>
            <p class="subtitle">Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Date Range: {date_range}</p>
        </div>

        <!-- Filter Bar -->
        <div style="background: linear-gradient(135deg, rgba(220, 38, 38, 0.15) 0%, rgba(153, 27, 27, 0.1) 100%); padding: 20px 30px; border-bottom: 2px solid rgba(220, 38, 38, 0.3);">
            <div style="max-width: 1600px; margin: 0 auto; display: flex; gap: 25px; align-items: center; flex-wrap: wrap;">
                <div style="display: flex; align-items: center; gap: 12px;">
                    <span style="color: #dc2626; font-weight: 600; font-size: 1.1rem; text-transform: uppercase; letter-spacing: 0.5px;">📊 View:</span>
                    <select id="filterType" style="padding: 10px 20px; border-radius: 10px; border: 2px solid #dc2626; background: linear-gradient(135deg, rgba(220, 38, 38, 0.2) 0%, rgba(153, 27, 27, 0.15) 100%); color: white; font-size: 1rem; font-weight: 600; cursor: pointer; min-width: 150px; transition: all 0.3s ease;">
                        <option value="all">All Data</option>
                        <option value="month">Monthly View</option>
                        <option value="week">Weekly View</option>
                    </select>
                </div>

                <div id="monthFilterContainer" style="display: none; flex: 1; min-width: 200px;">
                    <div style="display: flex; align-items: center; gap: 12px;">
                        <span style="color: rgba(255, 255, 255, 0.9); font-weight: 500; font-size: 0.95rem;">Select Month:</span>
                        <select id="monthFilter" style="flex: 1; padding: 10px 20px; border-radius: 10px; border: 2px solid #dc2626; background: rgba(255, 255, 255, 0.08); color: white; font-size: 1rem; font-weight: 500; cursor: pointer; transition: all 0.3s ease;">
                            {"".join(f'<option value="{month}">{month}</option>' for month in available_months)}
                        </select>
                    </div>
                </div>

                <div id="weekFilterContainer" style="display: none; flex: 1; min-width: 200px;">
                    <div style="display: flex; align-items: center; gap: 12px;">
                        <span style="color: rgba(255, 255, 255, 0.9); font-weight: 500; font-size: 0.95rem;">Select Week:</span>
                        <select id="weekFilter" style="flex: 1; padding: 10px 20px; border-radius: 10px; border: 2px solid #dc2626; background: rgba(255, 255, 255, 0.08); color: white; font-size: 1rem; font-weight: 500; cursor: pointer; transition: all 0.3s ease;">
                            <!-- Populated dynamically -->
                        </select>
                    </div>
                </div>
            </div>
        </div>

        <div class="summary-cards" id="summaryCards">
            <div class="summary-card production">
                <h3>Production Hours</h3>
                <div class="value">{total_production_hours:.2f}h</div>
                <div class="label">Excluding samples (480 min cycle)</div>
            </div>
            <div class="summary-card sample">
                <h3>Sample Hours</h3>
                <div class="value">{total_sample_hours:.2f}h</div>
                <div class="label">480 min cycle time only</div>
            </div>
            <div class="summary-card total">
                <h3>Total Hours</h3>
                <div class="value">{total_all_hours:.2f}h</div>
                <div class="label">All production time</div>
            </div>
            <div class="summary-card parts">
                <h3>Total Machines</h3>
                <div class="value">{total_machines}</div>
                <div class="label">{total_production_parts:,} production parts</div>
            </div>
            <div class="summary-card" style="--card-color: #f59e0b;">
                <h3>⏱️ Downtime</h3>
                <div class="value">{total_downtime_hours:.1f}h</div>
                <div class="label">{overall_downtime_percent:.1f}% of capacity</div>
            </div>
            <div class="summary-card" style="--card-color: #10b981;">
                <h3>✅ Availability</h3>
                <div class="value">{overall_availability:.1f}%</div>
                <div class="label">{total_all_hours:.1f}h / {total_available_hours:.1f}h capacity</div>
            </div>
        </div>

        <div class="content-section">
            <h2 class="section-title">📊 Machine Running Hours Comparison</h2>
            <div class="chart-container">
                <div class="chart-wrapper">
                    <canvas id="machineHoursChart"></canvas>
                </div>
            </div>
        </div>

        <div class="content-section">
            <h2 class="section-title">🔧 Machine Details Breakdown</h2>
            <select id="machineSelector" class="machine-selector">
                <option value="">Select a machine to view detailed breakdown...</option>
                {"".join(f'<option value="{i}">{machine}</option>' for i, (machine, _) in enumerate(sorted_machines))}
            </select>
            <div id="machineDetailsContainer"></div>
        </div>

        <div class="content-section">
            <h2 class="section-title">⚡ Working Time vs Downtime by Machine</h2>
            <div class="chart-container">
                <div class="chart-wrapper">
                    <canvas id="downtimeChart"></canvas>
                </div>
            </div>
        </div>

        <div class="content-section">
            <h2 class="section-title">🔄 Production vs Sample Hours</h2>
            <div class="chart-container">
                <div class="chart-wrapper">
                    <canvas id="hoursComparisonChart"></canvas>
                </div>
            </div>
        </div>

        <div class="footer">
            <p>⚠️ Note: Sample parts (480 minutes / 8 hours) counted ONCE per machine, not multiplied by parts</p>
            <p>Running Time Formula: (OK Parts × Cycle Time in minutes × 60) / 3600 = Hours</p>
        </div>
    </div>

    <script>
        // Global chart instances
        let machineHoursChartInstance, downtimeChartInstance, comparisonChartInstance;

        // Machine data
        const machineData = {json.dumps({
            machine: {
                'name': machine,
                'production_hours': stats['total_hours'],
                'sample_hours': stats['sample_hours'],
                'total_hours': stats['total_hours_with_samples'],
                'production_parts': stats['total_parts'],
                'sample_parts': stats['sample_parts'],
                'total_parts': stats['total_parts_with_samples'],
                'dates': stats['dates'],
                'shifts': stats['shifts'],
                'items': stats['items'],
                'available_capacity': stats['available_capacity'],
                'downtime_hours': stats['downtime_hours'],
                'downtime_percent': stats['downtime_percent'],
                'availability_percent': stats['availability_percent'],
                'shift_counts': dict(stats['shift_counts']),
                'records': [{
                    'date': r['date'],
                    'shift': r['shift'],
                    'item': r['item'],
                    'operation': r['operation'],
                    'ok_parts': r['ok_parts'],
                    'cycle_time': r['cycle_time'],
                    'is_sample': r['is_sample'],
                    'operator': r['operator']
                } for r in stats['records']]
            } for machine, stats in machine_stats.items()
        })};

        // Chart: Machine Hours Comparison
        const ctx = document.getElementById('machineHoursChart').getContext('2d');
        machineHoursChartInstance = new Chart(ctx, {{
            type: 'bar',
            data: {{
                labels: {json.dumps(machine_names)},
                datasets: [
                    {{
                        label: 'Production Hours',
                        data: {json.dumps(production_hours)},
                        backgroundColor: 'rgba(16, 185, 129, 0.8)',
                        borderColor: 'rgba(16, 185, 129, 1)',
                        borderWidth: 2
                    }},
                    {{
                        label: 'Sample Hours (480 min cycle)',
                        data: {json.dumps(sample_hours)},
                        backgroundColor: 'rgba(245, 158, 11, 0.8)',
                        borderColor: 'rgba(245, 158, 11, 1)',
                        borderWidth: 2
                    }}
                ]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{
                        labels: {{
                            color: 'white',
                            font: {{ size: 14 }}
                        }}
                    }},
                    tooltip: {{
                        backgroundColor: 'rgba(0, 0, 0, 0.8)',
                        titleColor: 'white',
                        bodyColor: 'white',
                        callbacks: {{
                            afterBody: function(context) {{
                                const index = context[0].dataIndex;
                                const machine = Object.keys(machineData)[index];
                                const data = machineData[machine];
                                return [
                                    '',
                                    'Total: ' + data.total_hours.toFixed(2) + 'h',
                                    'Production Parts: ' + data.production_parts.toLocaleString(),
                                    'Sample Parts: ' + data.sample_parts.toLocaleString()
                                ];
                            }}
                        }}
                    }}
                }},
                scales: {{
                    x: {{
                        stacked: true,
                        ticks: {{
                            color: 'white',
                            maxRotation: 45,
                            minRotation: 45
                        }},
                        grid: {{
                            color: 'rgba(255, 255, 255, 0.1)'
                        }}
                    }},
                    y: {{
                        stacked: true,
                        ticks: {{
                            color: 'white',
                            callback: function(value) {{
                                return value + 'h';
                            }}
                        }},
                        grid: {{
                            color: 'rgba(255, 255, 255, 0.1)'
                        }},
                        title: {{
                            display: true,
                            text: 'Running Hours',
                            color: 'white',
                            font: {{ size: 14 }}
                        }}
                    }}
                }}
            }}
        }});

        // Machine selector
        document.getElementById('machineSelector').addEventListener('change', function() {{
            const index = parseInt(this.value);
            const container = document.getElementById('machineDetailsContainer');

            if (isNaN(index)) {{
                container.innerHTML = '';
                return;
            }}

            const machineName = Object.keys(machineData)[index];
            const data = machineData[machineName];

            let html = `
                <div class="machine-details">
                    <h3>${{machineName}}</h3>
                    <div class="detail-grid">
                        <div class="detail-item">
                            <div class="label">Production Hours</div>
                            <div class="value">${{data.production_hours.toFixed(2)}}h</div>
                        </div>
                        <div class="detail-item">
                            <div class="label">Sample Hours</div>
                            <div class="value">${{data.sample_hours.toFixed(2)}}h</div>
                        </div>
                        <div class="detail-item">
                            <div class="label">Total Hours</div>
                            <div class="value">${{data.total_hours.toFixed(2)}}h</div>
                        </div>
                        <div class="detail-item">
                            <div class="label">Production Parts</div>
                            <div class="value">${{data.production_parts.toLocaleString()}}</div>
                        </div>
                        <div class="detail-item">
                            <div class="label">Sample Parts</div>
                            <div class="value">${{data.sample_parts.toLocaleString()}}</div>
                        </div>
                        <div class="detail-item">
                            <div class="label">Items Produced</div>
                            <div class="value">${{data.items.length}}</div>
                        </div>
                    </div>

                    <div style="margin-bottom: 15px;">
                        <strong>Active Dates:</strong> ${{data.dates.join(', ')}}
                    </div>
                    <div style="margin-bottom: 15px;">
                        <strong>Shifts:</strong> ${{data.shifts.join(', ')}}
                    </div>
                    <div style="margin-bottom: 15px;">
                        <strong>Items:</strong> ${{data.items.join(', ')}}
                    </div>

                    <h4 style="margin-top: 20px; margin-bottom: 15px; color: #dc2626; font-size: 1.2rem;">⚡ Downtime Analysis</h4>

                    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-bottom: 20px;">
                        <div style="background: linear-gradient(135deg, rgba(220, 38, 38, 0.2) 0%, rgba(153, 27, 27, 0.1) 100%); padding: 20px; border-radius: 12px; border: 1px solid rgba(220, 38, 38, 0.3); text-align: center;">
                            <div style="font-size: 0.85rem; color: rgba(255, 255, 255, 0.7); margin-bottom: 8px; text-transform: uppercase; font-weight: 600;">Available Capacity</div>
                            <div style="font-size: 2rem; font-weight: 700; color: #ffffff;">${{data.available_capacity.toFixed(1)}}h</div>
                            <div style="font-size: 0.8rem; color: rgba(255, 255, 255, 0.6); margin-top: 5px;">Total Shifts × 8h</div>
                        </div>

                        <div style="background: linear-gradient(135deg, rgba(5, 150, 105, 0.2) 0%, rgba(4, 120, 87, 0.1) 100%); padding: 20px; border-radius: 12px; border: 1px solid rgba(5, 150, 105, 0.3); text-align: center;">
                            <div style="font-size: 0.85rem; color: rgba(255, 255, 255, 0.7); margin-bottom: 8px; text-transform: uppercase; font-weight: 600;">Running Time</div>
                            <div style="font-size: 2rem; font-weight: 700; color: #10b981;">${{data.total_hours.toFixed(1)}}h</div>
                            <div style="font-size: 0.8rem; color: rgba(255, 255, 255, 0.6); margin-top: 5px;">Actual Production</div>
                        </div>

                        <div style="background: linear-gradient(135deg, rgba(245, 158, 11, 0.2) 0%, rgba(217, 119, 6, 0.1) 100%); padding: 20px; border-radius: 12px; border: 1px solid rgba(245, 158, 11, 0.3); text-align: center;">
                            <div style="font-size: 0.85rem; color: rgba(255, 255, 255, 0.7); margin-bottom: 8px; text-transform: uppercase; font-weight: 600;">Downtime</div>
                            <div style="font-size: 2rem; font-weight: 700; color: #f59e0b;">${{data.downtime_hours.toFixed(1)}}h</div>
                            <div style="font-size: 0.8rem; color: rgba(255, 255, 255, 0.6); margin-top: 5px;">${{data.downtime_percent.toFixed(1)}}% Lost</div>
                        </div>

                        <div style="background: linear-gradient(135deg, rgba(59, 130, 246, 0.2) 0%, rgba(37, 99, 235, 0.1) 100%); padding: 20px; border-radius: 12px; border: 1px solid rgba(59, 130, 246, 0.3); text-align: center;">
                            <div style="font-size: 0.85rem; color: rgba(255, 255, 255, 0.7); margin-bottom: 8px; text-transform: uppercase; font-weight: 600;">Availability</div>
                            <div style="font-size: 2rem; font-weight: 700; color: #3b82f6;">${{data.availability_percent.toFixed(1)}}%</div>
                            <div style="font-size: 0.8rem; color: rgba(255, 255, 255, 0.6); margin-top: 5px;">Machine Uptime</div>
                        </div>
                    </div>

                    <div style="background: rgba(255, 255, 255, 0.05); padding: 20px; border-radius: 12px; border-left: 4px solid #dc2626;">
                        <h5 style="color: #dc2626; font-size: 1rem; margin-bottom: 15px;">📊 Shift Breakdown</h5>
                        <div style="display: flex; flex-direction: column; gap: 12px;">
                            ${{Object.entries(data.shift_counts).map(([shift, count]) => `
                                <div style="display: flex; justify-content: space-between; align-items: center; padding: 12px; background: rgba(0,0,0,0.2); border-radius: 8px;">
                                    <span style="font-weight: 600; color: #ffffff; font-size: 1rem;">${{shift}}</span>
                                    <div style="text-align: right;">
                                        <div style="font-size: 1.2rem; font-weight: 700; color: #dc2626;">${{count}} runs</div>
                                        <div style="font-size: 0.85rem; color: rgba(255,255,255,0.6);">${{(count * 8).toFixed(1)}}h capacity</div>
                                    </div>
                                </div>
                            `).join('')}}
                        </div>
                    </div>
                </div>
            `;

            container.innerHTML = html;
        }});

        // Chart: Working Time vs Downtime by Machine
        const downtimeCtx = document.getElementById('downtimeChart').getContext('2d');
        const machineDowntimeData = Object.entries(machineData).sort((a, b) => b[1].total_hours - a[1].total_hours);
        downtimeChartInstance = new Chart(downtimeCtx, {{
            type: 'bar',
            data: {{
                labels: machineDowntimeData.map(([name, _]) => name),
                datasets: [
                    {{
                        label: 'Working Time',
                        data: machineDowntimeData.map(([_, data]) => data.total_hours),
                        backgroundColor: 'rgba(16, 185, 129, 0.8)',
                        borderColor: 'rgba(16, 185, 129, 1)',
                        borderWidth: 2
                    }},
                    {{
                        label: 'Downtime',
                        data: machineDowntimeData.map(([_, data]) => data.downtime_hours),
                        backgroundColor: 'rgba(245, 158, 11, 0.8)',
                        borderColor: 'rgba(245, 158, 11, 1)',
                        borderWidth: 2
                    }}
                ]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{
                        display: true,
                        labels: {{ color: 'white', font: {{ size: 14 }} }}
                    }},
                    title: {{
                        display: true,
                        text: 'Working Hours vs Downtime per Machine',
                        color: 'white',
                        font: {{ size: 16 }}
                    }},
                    tooltip: {{
                        callbacks: {{
                            afterBody: function(context) {{
                                const index = context[0].dataIndex;
                                const [machineName, data] = machineDowntimeData[index];
                                return [
                                    '',
                                    'Available Capacity: ' + data.available_capacity.toFixed(1) + 'h',
                                    'Availability: ' + data.availability_percent.toFixed(1) + '%',
                                    'Downtime: ' + data.downtime_percent.toFixed(1) + '%'
                                ];
                            }}
                        }}
                    }}
                }},
                scales: {{
                    x: {{
                        stacked: true,
                        ticks: {{
                            color: 'white',
                            maxRotation: 45,
                            minRotation: 45
                        }},
                        grid: {{ color: 'rgba(255, 255, 255, 0.1)' }}
                    }},
                    y: {{
                        stacked: true,
                        ticks: {{
                            color: 'white',
                            callback: function(value) {{ return value.toFixed(1) + 'h'; }}
                        }},
                        grid: {{ color: 'rgba(255, 255, 255, 0.1)' }},
                        title: {{
                            display: true,
                            text: 'Hours',
                            color: 'white',
                            font: {{ size: 14 }}
                        }}
                    }}
                }}
            }}
        }});


        // Chart: Production vs Sample Hours Comparison
        const comparisonCtx = document.getElementById('hoursComparisonChart').getContext('2d');
        comparisonChartInstance = new Chart(comparisonCtx, {{
            type: 'doughnut',
            data: {{
                labels: ['Production Hours', 'Sample Hours (480 min once per machine)'],
                datasets: [{{
                    data: [
                        {total_production_hours:.2f},
                        {total_sample_hours:.2f}
                    ],
                    backgroundColor: [
                        'rgba(5, 150, 105, 0.8)',
                        'rgba(245, 158, 11, 0.8)'
                    ],
                    borderColor: [
                        'rgba(5, 150, 105, 1)',
                        'rgba(245, 158, 11, 1)'
                    ],
                    borderWidth: 2
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{
                        display: true,
                        position: 'bottom',
                        labels: {{ color: 'white', font: {{ size: 14 }}, padding: 20 }}
                    }},
                    title: {{
                        display: true,
                        text: 'Total Hours Breakdown',
                        color: 'white',
                        font: {{ size: 16 }}
                    }},
                    tooltip: {{
                        callbacks: {{
                            label: function(context) {{
                                const label = context.label || '';
                                const value = context.parsed || 0;
                                const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                const percentage = ((value / total) * 100).toFixed(1);
                                return `${{label}}: ${{value.toFixed(2)}}h (${{percentage}}%)`;
                            }}
                        }}
                    }}
                }}
            }}
        }});

        // ============================================================================
        // FILTERING FUNCTIONALITY - Must come AFTER all charts are created
        // ============================================================================

        // Helper functions for date/week handling
        function getWeekNumber(dateStr) {{
            const date = new Date(dateStr);
            const firstDayOfYear = new Date(date.getFullYear(), 0, 1);
            const pastDaysOfYear = (date - firstDayOfYear) / 86400000;
            return Math.ceil((pastDaysOfYear + firstDayOfYear.getDay() + 1) / 7);
        }}

        function getWeekLabel(dateStr) {{
            const date = new Date(dateStr);
            const year = date.getFullYear();
            const week = getWeekNumber(dateStr);
            return `${{year}}-W${{String(week).padStart(2, '0')}}`;
        }}

        // Populate week filter with available weeks from data
        function populateWeekFilter() {{
            const weeks = new Set();
            Object.values(machineData).forEach(machine => {{
                machine.dates.forEach(date => {{
                    if (date && date !== 'Unknown') {{
                        weeks.add(getWeekLabel(date));
                    }}
                }});
            }});

            const weekFilter = document.getElementById('weekFilter');
            weekFilter.innerHTML = '';
            const sortedWeeks = Array.from(weeks).sort();
            sortedWeeks.forEach(week => {{
                const option = document.createElement('option');
                option.value = week;
                option.textContent = week;
                weekFilter.appendChild(option);
            }});
        }}

        // Update summary cards with filtered data
        function updateSummaryCards(filteredData) {{
            const totalProductionHours = Object.values(filteredData).reduce((sum, d) => sum + d.total_hours, 0);
            const totalSampleHours = Object.values(filteredData).reduce((sum, d) => sum + d.sample_hours, 0);
            const totalAllHours = Object.values(filteredData).reduce((sum, d) => sum + d.total_hours_with_samples, 0);
            const totalDowntimeHours = Object.values(filteredData).reduce((sum, d) => sum + d.downtime_hours, 0);
            const totalAvailableHours = Object.values(filteredData).reduce((sum, d) => sum + d.available_capacity, 0);
            const overallAvailability = totalAvailableHours > 0 ? (totalAllHours / totalAvailableHours * 100) : 0;
            const overallDowntimePercent = 100 - overallAvailability;
            const totalMachines = Object.keys(filteredData).length;
            const totalProductionParts = Object.values(filteredData).reduce((sum, d) => sum + d.production_parts, 0);

            // Update card values
            document.querySelector('.summary-card.production .value').textContent = totalProductionHours.toFixed(2) + 'h';
            document.querySelector('.summary-card.sample .value').textContent = totalSampleHours.toFixed(2) + 'h';
            document.querySelector('.summary-card.total .value').textContent = totalAllHours.toFixed(2) + 'h';
            document.querySelector('.summary-card.parts .value').textContent = totalMachines;
            document.querySelector('.summary-card.parts .label').textContent = totalProductionParts.toLocaleString() + ' production parts';

            // Update downtime card
            const downtimeCard = Array.from(document.querySelectorAll('.summary-card')).find(c => c.innerHTML.includes('Downtime'));
            if (downtimeCard) {{
                downtimeCard.querySelector('.value').textContent = totalDowntimeHours.toFixed(1) + 'h';
                downtimeCard.querySelector('.label').textContent = overallDowntimePercent.toFixed(1) + '% of capacity';
            }}

            // Update availability card
            const availabilityCard = Array.from(document.querySelectorAll('.summary-card')).find(c => c.innerHTML.includes('Availability'));
            if (availabilityCard) {{
                availabilityCard.querySelector('.value').textContent = overallAvailability.toFixed(1) + '%';
                availabilityCard.querySelector('.label').textContent = totalAllHours.toFixed(1) + 'h / ' + totalAvailableHours.toFixed(1) + 'h capacity';
            }}
        }}

        // Update all charts with filtered data
        function updateCharts(filteredData) {{
            const sortedByHours = Object.entries(filteredData).sort((a, b) => b[1].total_hours - a[1].total_hours);

            // Update machine hours chart
            machineHoursChartInstance.data.labels = sortedByHours.map(([name, _]) => name);
            machineHoursChartInstance.data.datasets[0].data = sortedByHours.map(([_, d]) => d.total_hours);
            machineHoursChartInstance.data.datasets[1].data = sortedByHours.map(([_, d]) => d.sample_hours);
            machineHoursChartInstance.update();

            // Update downtime chart
            downtimeChartInstance.data.labels = sortedByHours.map(([name, _]) => name);
            downtimeChartInstance.data.datasets[0].data = sortedByHours.map(([_, d]) => d.total_hours_with_samples);
            downtimeChartInstance.data.datasets[1].data = sortedByHours.map(([_, d]) => d.downtime_hours);
            downtimeChartInstance.update();

            // Update comparison chart
            const totalProductionHours = Object.values(filteredData).reduce((sum, d) => sum + d.total_hours, 0);
            const totalSampleHours = Object.values(filteredData).reduce((sum, d) => sum + d.sample_hours, 0);
            comparisonChartInstance.data.datasets[0].data = [totalProductionHours, totalSampleHours];
            comparisonChartInstance.update();
        }}

        // Apply filter to data and recalculate statistics
        function applyFilter(filterType, filterValue) {{
            console.log('Applying filter:', filterType, filterValue);

            // If "all" is selected, use original data
            if (filterType === 'all') {{
                const originalData = {{}};
                Object.entries(machineData).forEach(([machineName, data]) => {{
                    originalData[machineName] = {{
                        total_hours: data.total_hours,
                        sample_hours: data.sample_hours,
                        total_hours_with_samples: data.total_hours_with_samples,
                        production_parts: data.production_parts,
                        sample_parts: data.sample_parts,
                        available_capacity: data.available_capacity,
                        downtime_hours: data.downtime_hours,
                        availability_percent: data.availability_percent,
                        downtime_percent: data.downtime_percent
                    }};
                }});
                updateSummaryCards(originalData);
                updateCharts(originalData);
                return;
            }}

            // Filter machine data based on selection
            let filteredData = {{}};

            Object.entries(machineData).forEach(([machineName, data]) => {{
                let filteredRecords = data.records;

                if (filterType === 'month' && filterValue) {{
                    filteredRecords = data.records.filter(r => r.date && r.date.startsWith(filterValue));
                }} else if (filterType === 'week' && filterValue) {{
                    filteredRecords = data.records.filter(r => r.date && getWeekLabel(r.date) === filterValue);
                }}

                if (filteredRecords.length > 0) {{
                    // Recalculate statistics for filtered records
                    let totalSeconds = 0;
                    let sampleSeconds = 0;
                    let totalParts = 0;
                    let sampleParts = 0;
                    let hasSamples = false;
                    let shiftCounts = {{}};

                    filteredRecords.forEach(record => {{
                        // Count shift occurrences
                        shiftCounts[record.shift] = (shiftCounts[record.shift] || 0) + 1;

                        if (record.is_sample) {{
                            if (!hasSamples) {{
                                sampleSeconds = 28800; // 480 min = 8 hours = 28,800 sec
                                hasSamples = true;
                            }}
                            sampleParts += record.ok_parts;
                        }} else {{
                            totalSeconds += record.ok_parts * record.cycle_time;
                            totalParts += record.ok_parts;
                        }}
                    }});

                    const totalHours = totalSeconds / 3600;
                    const sampleHours = sampleSeconds / 3600;
                    const totalHoursWithSamples = totalHours + sampleHours;

                    // Calculate downtime
                    const totalShiftOccurrences = Object.values(shiftCounts).reduce((a, b) => a + b, 0);
                    const availableCapacity = totalShiftOccurrences * 8; // 8 hours per shift
                    const downtimeHours = Math.max(0, availableCapacity - totalHoursWithSamples);
                    const availabilityPercent = availableCapacity > 0 ? (totalHoursWithSamples / availableCapacity * 100) : 0;
                    const downtimePercent = 100 - availabilityPercent;

                    filteredData[machineName] = {{
                        total_hours: totalHours,
                        sample_hours: sampleHours,
                        total_hours_with_samples: totalHoursWithSamples,
                        production_parts: totalParts,
                        sample_parts: sampleParts,
                        available_capacity: availableCapacity,
                        downtime_hours: downtimeHours,
                        availability_percent: availabilityPercent,
                        downtime_percent: downtimePercent
                    }};
                }}
            }});

            // Update UI with filtered data
            updateSummaryCards(filteredData);
            updateCharts(filteredData);
        }}

        // ============================================================================
        // INITIALIZE FILTERS - Set up event listeners
        // ============================================================================

        // Populate week dropdown
        populateWeekFilter();

        // Filter type change handler
        document.getElementById('filterType').addEventListener('change', function() {{
            const filterType = this.value;
            document.getElementById('monthFilterContainer').style.display = filterType === 'month' ? 'flex' : 'none';
            document.getElementById('weekFilterContainer').style.display = filterType === 'week' ? 'flex' : 'none';

            if (filterType === 'all') {{
                applyFilter('all', null);
            }} else if (filterType === 'month') {{
                applyFilter('month', document.getElementById('monthFilter').value);
            }} else if (filterType === 'week') {{
                applyFilter('week', document.getElementById('weekFilter').value);
            }}
        }});

        // Month filter change handler
        document.getElementById('monthFilter').addEventListener('change', function() {{
            applyFilter('month', this.value);
        }});

        // Week filter change handler
        document.getElementById('weekFilter').addEventListener('change', function() {{
            applyFilter('week', this.value);
        }});

        console.log('✅ Dashboard initialized with filtering functionality');
    </script>
</body>
</html>"""

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"✅ HTML report saved to: {output_file}")


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
    report_lines.append(f"Total Sample Hours (cycle time = 480 min): {total_sample_hours:.2f} hours")
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
        report_lines.append(f"  Sample Hours (480 min cycle time):  {stats['sample_hours']:.2f} hours")
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

    # Footer
    report_lines.append("=" * 100)
    report_lines.append("NOTE: Sample parts (480 minutes / 8 hours) counted ONCE per machine, not multiplied by parts")
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


def select_folder_dialog(title="Select Folder", prompt_text="Enter folder path: "):
    """Show GUI dialog to select a folder, or fall back to text input"""
    if TKINTER_AVAILABLE:
        try:
            root = tk.Tk()
            root.withdraw()  # Hide the main window
            root.attributes('-topmost', True)  # Bring dialog to front
            folder_path = filedialog.askdirectory(title=title)
            root.destroy()
            return folder_path
        except Exception as e:
            print(f"⚠️  GUI dialog failed: {e}")
            print("   Falling back to text input...")

    # Fallback to text input
    return input(prompt_text).strip()


def get_available_months(records):
    """Extract unique months from records"""
    months = set()
    for record in records:
        date_str = record.get('date', '')
        if date_str and date_str != 'Unknown':
            # Extract YYYY-MM from date
            try:
                month = date_str[:7]  # Get YYYY-MM
                months.add(month)
            except:
                pass
    return sorted(list(months))


def filter_records_by_month(records, selected_months):
    """Filter records to only include selected months"""
    if not selected_months or 'all' in [m.lower() for m in selected_months]:
        return records

    filtered = []
    for record in records:
        date_str = record.get('date', '')
        if date_str and date_str != 'Unknown':
            month = date_str[:7]  # Get YYYY-MM
            if month in selected_months:
                filtered.append(record)

    return filtered


def main():
    """Main function to analyze CNC machine running times"""
    print("🏭 CNC Machine Running Time Analysis")
    print("=" * 70)
    print()

    # Get input folder path from command line or GUI dialog
    if len(sys.argv) > 1:
        folder_path = sys.argv[1]
    else:
        print("📂 INPUT FOLDER - Please select the folder containing HTML report files...")
        folder_path = select_folder_dialog(
            title="Select INPUT Folder (HTML Reports)",
            prompt_text="Enter the folder path containing HTML reports: "
        )

        if not folder_path:
            print("❌ No folder selected. Exiting...")
            return

    if not os.path.exists(folder_path):
        print(f"❌ Error: Folder not found: {folder_path}")
        return

    print(f"✅ Input folder: {folder_path}")
    print()

    # Find all HTML files
    html_files = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.html') and not file.startswith('cnc_running_hours_dashboard'):
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

    # Show available months (info only, no prompting)
    available_months = get_available_months(all_records)
    if available_months:
        print("📅 Available months in data:")
        for i, month in enumerate(available_months, 1):
            month_records = [r for r in all_records if r.get('date', '').startswith(month)]
            print(f"   {i}. {month} ({len(month_records)} records)")
        print()
        print("✅ Processing ALL data - filtering available in dashboard")
    else:
        print("⚠️  No date information found in records")
    print()

    # Calculate running hours from ALL records
    machine_stats = calculate_running_hours(all_records, exclude_samples=True)

    # Get output folder path from command line or GUI dialog
    print()
    if len(sys.argv) > 2:
        output_folder = sys.argv[2]
    else:
        print("💾 OUTPUT FOLDER - Please select where to save the reports...")
        output_folder = select_folder_dialog(
            title="Select OUTPUT Folder (Save Reports)",
            prompt_text="Enter the output folder path (or press Enter for current directory): "
        )

        if not output_folder:
            print("⚠️  No output folder selected, using current directory...")
            output_folder = "."

    # Create output folder if it doesn't exist
    if output_folder != ".":
        os.makedirs(output_folder, exist_ok=True)

    if not os.path.exists(output_folder):
        print(f"❌ Error: Could not create output folder: {output_folder}")
        return

    print(f"✅ Output folder: {os.path.abspath(output_folder)}")
    print()

    # Generate reports
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Generate HTML report (main output)
    html_file = os.path.join(output_folder, f"cnc_running_hours_dashboard_{timestamp}.html")
    generate_html_report(machine_stats, all_records, html_file, available_months)

    # Generate text report (backup)
    txt_file = os.path.join(output_folder, f"cnc_running_hours_report_{timestamp}.txt")
    generate_report(machine_stats, txt_file)

    # Also save JSON data for further analysis
    json_file = os.path.join(output_folder, f"cnc_running_hours_data_{timestamp}.json")

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
    print("🎉 Analysis complete!")
    print(f"📊 Open the HTML dashboard: {html_file}")


if __name__ == "__main__":
    main()
