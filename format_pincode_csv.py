#!/usr/bin/env python3
"""
Format All India Pincode Directory CSV for Google Sheets RawImport tab.

Input: CSV downloaded from data.gov.in or similar source.
Output: RawImport_ready.csv with 4 columns:
        State/UT, District, Pincode, Tehsil/Block
"""
import sys
import csv

def format_pincode_csv(input_filepath, output_filepath="RawImport_ready.csv"):
    print(f"Processing {input_filepath}...")
    
    # Common column header variations in data.gov.in datasets
    state_keys = ['state', 'statename', 'state_name', 'state/ut']
    district_keys = ['district', 'districtname', 'district_name']
    pincode_keys = ['pincode', 'pin', 'pincode/pin code', 'pin_code']
    tehsil_keys = ['taluk', 'tehsil', 'block', 'sub_district', 'officename', 'office_name']

    rows_out = []
    
    with open(input_filepath, 'r', encoding='utf-8-sig', errors='replace') as infile:
        reader = csv.DictReader(infile)
        fieldnames = [f.strip().lower() for f in reader.fieldnames or []]
        field_map = {f.strip().lower(): f for f in reader.fieldnames or []}

        # Find matching columns
        state_col = next((field_map[f] for f in fieldnames if any(k in f for k in state_keys)), None)
        district_col = next((field_map[f] for f in fieldnames if any(k in f for k in district_keys)), None)
        pincode_col = next((field_map[f] for f in fieldnames if any(k in f for k in pincode_keys)), None)
        tehsil_col = next((field_map[f] for f in fieldnames if any(k in f for k in tehsil_keys)), None)

        print(f"Detected Columns -> State: '{state_col}', District: '{district_col}', Pincode: '{pincode_col}', Tehsil: '{tehsil_col}'")

        for row in reader:
            state = row.get(state_col, '').strip().title() if state_col else ''
            district = row.get(district_col, '').strip().title() if district_col else ''
            pincode = row.get(pincode_col, '').strip() if pincode_col else ''
            tehsil = row.get(tehsil_col, '').strip().title() if tehsil_col else ''

            if state and district and pincode:
                rows_out.append([state, district, pincode, tehsil])

    with open(output_filepath, 'w', encoding='utf-8', newline='') as outfile:
        writer = csv.writer(outfile)
        writer.writerow(['State/UT', 'District', 'Pincode', 'Tehsil/Block'])
        writer.writerows(rows_out)

    print(f"✅ Success: Saved {len(rows_out)} formatted rows to {output_filepath}")

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python3 format_pincode_csv.py <path_to_raw_pincode_data.csv>")
        sys.exit(1)
    format_pincode_csv(sys.argv[1])
