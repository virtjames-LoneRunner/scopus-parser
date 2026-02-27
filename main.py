import os
import argparse

from scopus_parser.scopus import BISUScopus
from scopus_parser.config import BISUScopusConfig

def main():
    parser = argparse.ArgumentParser(description="Welcome to the BISU Scopus API Client!")
    parser.add_argument("--keywords", nargs="+", help="Keywords to search for.")
    parser.add_argument("--source_types", nargs="+", help="Filter by Source Type: (Journal/Book Series/Trade Journal/etc.)")
    parser.add_argument("--active_status", nargs="+", help="Filter by Activity: (Active/Inactive)")
    parser.add_argument("--output_filename", type=str, help="Output filename (must end in .xlsx)")
    args = parser.parse_args()

    config = BISUScopusConfig()
    config.list_link = "https://downloads.ctfassets.net/o78em1y1w4i4/7xtaTxNiNcWRTeZkV86eNy/de9e757c475827b03206a5bf4d24c8a3/ext_list_Jan_2026.xlsx"
    config.list_filename = os.path.join("scopus", "ext_list_Jan_2026.xlsx")

    bisu_scopus = BISUScopus(config)
    try:
        print("Checking for list file.")
        list_file = open(config.list_filename, 'r')
    except FileNotFoundError:
        print("Scopus list file not found in scopus directory.")
        print("Trying to download file......")
        downloaded = bisu_scopus.retrieve_list()
        if not downloaded:
            print("Please download the scopus list first!")
            exit()

    try:
        bisu_scopus.load_file()
        bisu_scopus.read_sources()
    except Exception as e:
        print("An error occurred while trying to load the scopus list file.")
        print(e)
        exit()

    try:
        bisu_scopus.filter_by_column("Source Title", args.keywords)
        if args.source_types:
            bisu_scopus.filter_by_column("Source Type", args.source_types)
        if args.active_status:
            bisu_scopus.filter_by_column("Active or Inactive", args.active_status)

        # Apply the mapping
        field_descriptions = bisu_scopus.filtered['All Science Journal Classification Codes (ASJC)'].apply(bisu_scopus.map_multiple_codes)
        # Insert at a specific index (Index 1 for the 2nd column)
        bisu_scopus.filtered.insert(1, 'Field Descriptions', field_descriptions)
        
        bisu_scopus.add_scimago_rankings(os.path.join("scopus", "scimagojr 2024.csv"))
        bisu_scopus.clean_up_columns()
        
        # Export data to excel file
        export_filename = 'scopus_api_filtered_results.xlsx'
        if (args.output_filename):
            if ".xlsx" not in args.output_filename:
                export_filename = args.output_filename + ".xlsx"
            export_filename = args.output_filename
        bisu_scopus.save_with_formatting(export_filename)
        # bisu_scopus.export_filtered(export_filename)
            
    except Exception as e:
        print("An error was encountered while trying to filter using keywords.")
        print(e)
        exit()













if __name__ == "__main__":
    main()