import os
import argparse

from scopus_parser.scopus import BISUScopus
from scopus_parser.config import BISUScopusConfig

def main():
    parser = argparse.ArgumentParser(
        description="Filter Scopus sources for publication-submission shortlisting."
    )
    parser.add_argument("--keywords", nargs="+", help="Keywords to search for.", required=True)
    parser.add_argument(
        "--search_fields",
        nargs="+",
        choices=["title", "categories", "areas", "asjc", "field_descriptions", "publisher"],
        default=["title", "categories", "areas", "asjc", "field_descriptions"],
        help="Fields searched by keywords.",
    )
    parser.add_argument("--match_mode", choices=["any", "all"], default="any",
                        help="Match any keyword or require all keywords.")
    parser.add_argument("--language", nargs="+", help="Article Language in Source (Three-Letter ISO Language Codes) to search for. (ENG for English)")
    parser.add_argument("--source_types", nargs="+", help="Filter by Source Type: (Journal/Book Series/Trade Journal/etc.)")
    parser.add_argument("--active_status", nargs="+", help="Filter by Activity: (Active/Inactive)")
    parser.add_argument("--categories", nargs="+", help="Filter by SCImago category.")
    parser.add_argument("--areas", nargs="+", help="Filter by SCImago area.")
    parser.add_argument("--quartiles", nargs="+", choices=["Q1", "Q2", "Q3", "Q4"],
                        help="Keep journals in selected SJR quartiles.")
    parser.add_argument("--min_sjr", type=float, help="Minimum SJR score.")
    parser.add_argument("--min_h_index", type=int, help="Minimum H-index.")
    parser.add_argument("--sort_by", choices=["sjr", "h_index"], default="sjr",
                        help="Sort the final shortlist.")
    parser.add_argument("--exclude_discontinued", action="store_true",
                        help="Exclude titles discontinued by Scopus.")
    parser.add_argument("--scimago_list_file", type=str, help="Default: scimagojr 2025.csv")
    parser.add_argument("--output_filename", type=str, help="Default: scopus_api_filtered_results.xlsx")
    args = parser.parse_args()

    config = BISUScopusConfig()
    config.list_link = "https://downloads.ctfassets.net/o78em1y1w4i4/7xtaTxNiNcWRTeZkV86eNy/de9e757c475827b03206a5bf4d24c8a3/ext_list_Jan_2026.xlsx"
    config.list_filename = os.path.join("scopus", "ext_list_Jan_2026.xlsx")

    bisu_scopus = BISUScopus(config)
    print("Checking for list file.")
    if not os.path.isfile(config.list_filename):
        print("Scopus list file not found in scopus directory.")
        print("Trying to download file......")
        downloaded = bisu_scopus.retrieve_list()
        if not downloaded:
            print(f"Could not find or download {config.list_filename}.")
            return

    try:
        bisu_scopus.load_file()
        bisu_scopus.read_sources()
    except Exception as e:
        print("An error occurred while trying to load the scopus list file.")
        print(e)
        exit()

    try:
        if args.source_types:
            bisu_scopus.filter_by_values("Source Type", args.source_types)
        if args.active_status:
            bisu_scopus.filter_by_values("Active or Inactive", args.active_status)
        if args.language:
            bisu_scopus.filter_by_values(
                "Article Language in Source (Three-Letter ISO Language Codes)",
                args.language,
            )
        if args.exclude_discontinued:
            bisu_scopus.exclude_discontinued()

        target = bisu_scopus.filtered if bisu_scopus.filtered is not None else bisu_scopus.data_list_sources
        target = target.copy()
        field_descriptions = target['All Science Journal Classification Codes (ASJC)'].apply(
            bisu_scopus.map_multiple_codes
        )
        if 'Field Descriptions' in target.columns:
            target['Field Descriptions'] = field_descriptions
        else:
            target.insert(1, 'Field Descriptions', field_descriptions)
        bisu_scopus.filtered = target
        
        scimago_list_file = args.scimago_list_file or "scimagojr 2025.csv"
        scimago_path = os.path.join("scopus", scimago_list_file)
        if not os.path.isfile(scimago_path):
            raise FileNotFoundError(scimago_path)
        bisu_scopus.add_scimago_rankings(scimago_path)

        if args.categories:
            bisu_scopus.filter_by_values("Categories", args.categories)
        if args.areas:
            bisu_scopus.filter_by_values("Areas", args.areas)
        if args.quartiles:
            bisu_scopus.filter_by_values("SJR Best Quartile", args.quartiles)
        bisu_scopus.filter_by_minimum("SJR", args.min_sjr)
        bisu_scopus.filter_by_minimum("H index", args.min_h_index)

        search_column_map = {
            "title": "Source Title",
            "categories": "Categories",
            "areas": "Areas",
            "asjc": "All Science Journal Classification Codes (ASJC)",
            "field_descriptions": "Field Descriptions",
            "publisher": "Publisher",
        }
        search_columns = [search_column_map[field] for field in args.search_fields]
        bisu_scopus.filter_by_keywords(search_columns, args.keywords, args.match_mode)
        bisu_scopus.sort_shortlist(args.sort_by)

        bisu_scopus.clean_up_columns()
        bisu_scopus.print_filter_summary()
        
        # Export data to excel file
        export_filename = args.output_filename or 'scopus_api_filtered_results.xlsx'
        if not export_filename.lower().endswith(".xlsx"):
            export_filename += ".xlsx"
        bisu_scopus.save_with_formatting(export_filename)
        # bisu_scopus.export_filtered(export_filename)
            
    except Exception as e:
        print("An error was encountered while trying to filter using keywords.")
        print(e)
        exit()













if __name__ == "__main__":
    main()