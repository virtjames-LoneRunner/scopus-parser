import requests
import pandas
import os

from scopus_parser.config import BISUScopusConfig


class BISUScopus():
    def __init__(self, config: BISUScopusConfig):
        self.list_link = config.list_link
        self.list_filename = config.list_filename
        self.columns_to_keep = config.columns_to_keep
        self.list_file = None
        self.list_file_sheet_names = []
        self.data_list_sources = None
        self.data_asjc = None
        self.filtered = None
    
    def retrieve_list(self):
        if (self.list_link == ""):
            print("Please provide the link")
            return False
        os.makedirs(os.path.dirname(self.list_filename) or ".", exist_ok=True)
        try:
            response = requests.get(self.list_link, timeout=60)
            response.raise_for_status()
            with open(self.list_filename, 'wb') as f:
                f.write(response.content)
            print(f"Successfully written to {self.list_filename}.")
            return True
        except requests.RequestException as exc:
            print(f"There was an error while downloading the Scopus list: {exc}")
            return False

    def load_file(self):
        print("Loading list file....")
        self.list_file = pandas.ExcelFile(self.list_filename)
        self.list_file_sheet_names = self.list_file.sheet_names
    
    def read_sources(self):
        print("Reading scopus sources from list....")
        self.data_list_sources = pandas.read_excel(self.list_filename)

        mapping_sheet = next(
            (name for name in self.list_file_sheet_names if "ASJC" in name.upper()),
            self.list_file_sheet_names[-1],
        )
        df = pandas.read_excel(self.list_file, mapping_sheet)
        mapping_df = df.dropna(subset=df.columns[:2], how='any')

        # Clean up: Ensure the 'Code' column is numeric and Descriptions are strings
        # This ignores header rows like "Code" or "Description"
        mapping_df.columns = ['Code', 'Description'] + list(mapping_df.columns[2:])
        mapping_df['Code'] = pandas.to_numeric(mapping_df['Code'], errors='coerce')
        mapping_df = mapping_df.dropna(subset=['Code'])

        # Final Clean Table (Only first two columns)
        asjc_lookup_table = mapping_df[['Code', 'Description']].reset_index(drop=True)
        self.data_asjc = {
            int(code): description
            for code, description in zip(asjc_lookup_table['Code'], asjc_lookup_table['Description'])
            if pandas.notna(code)
        }

    def filter_by_column(self, column, keywords=None):
        """Backward-compatible substring filter for a single column."""
        return self.filter_by_values(column, keywords or [])

    def filter_by_values(self, column, values):
        """Keep rows containing any selected value in a column."""
        if not values:
            return
        if column not in (self.filtered if self.filtered is not None else self.data_list_sources).columns:
            raise KeyError(f"Unknown filter column: {column}")

        target = self.filtered if self.filtered is not None else self.data_list_sources
        wanted = [str(value).strip().casefold() for value in values if str(value).strip()]

        def matches(value):
            if pandas.isna(value):
                return False
            text = str(value).casefold()
            return any(wanted_value in text for wanted_value in wanted)

        mask = target[column].apply(matches)
        self.filtered = target.loc[mask].copy()

    def filter_by_keywords(self, columns, keywords, match_mode="any"):
        """Search case-insensitive keyword substrings across one or more columns."""
        if not keywords:
            return
        wanted = [str(keyword).strip().casefold() for keyword in keywords if str(keyword).strip()]
        if not wanted:
            return
        target = self.filtered if self.filtered is not None else self.data_list_sources

        missing = [column for column in columns if column not in target.columns]
        if missing:
            raise KeyError(f"Unknown search column(s): {', '.join(missing)}")

        def row_matches(row):
            matches = []
            for keyword in wanted:
                keyword_matches = any(
                    not pandas.isna(row[column])
                    and keyword in str(row[column]).casefold()
                    for column in columns
                )
                matches.append(keyword_matches)
            return all(matches) if match_mode == "all" else any(matches)

        mask = target.apply(row_matches, axis=1)
        self.filtered = target.loc[mask].copy()

    def filter_by_minimum(self, column, minimum):
        if minimum is None:
            return
        target = self.filtered if self.filtered is not None else self.data_list_sources
        if column not in target.columns:
            raise KeyError(f"Unknown ranking column: {column}")
        numeric = pandas.to_numeric(target[column], errors="coerce")
        self.filtered = target.loc[numeric >= minimum].copy()

    def exclude_discontinued(self):
        target = self.filtered if self.filtered is not None else self.data_list_sources
        column = "Titles Discontinued by Scopus"
        if column not in target.columns:
            raise KeyError(f"Unknown discontinued-title column: {column}")
        self.filtered = target.loc[target[column].isna()].copy()

    def sort_shortlist(self, sort_by="sjr"):
        target = self.filtered if self.filtered is not None else self.data_list_sources
        column = "SJR" if sort_by == "sjr" else "H index"
        if column not in target.columns:
            raise KeyError(f"Unknown sort column: {column}")
        numeric = pandas.to_numeric(target[column], errors="coerce")
        sorted_target = target.assign(**{f"__{column}": numeric}).sort_values(
            f"__{column}",
            ascending=False,
            na_position="last",
            kind="stable",
        )
        sorted_target = sorted_target.drop(columns=[f"__{column}"])
        self.filtered = sorted_target
    
    def map_multiple_codes(self, val):
        if pandas.isna(val):
            return None
        # Split by semicolon and strip whitespace
        codes = [c.strip() for c in str(val).split(';')]
        # Look up each code and join descriptions with a semicolon
        descriptions = [self.data_asjc.get(int(c), c) for c in codes if c.isdigit()]
        return "; ".join(descriptions)
    
    
    def add_scimago_rankings(self, scimago_csv_path):
        """Merges SCImago rankings into the filtered results."""
        print("Merging SCImago ranking data...")
        
        sjr_df = pandas.read_csv(scimago_csv_path, sep=';', encoding='utf-8-sig')

        sjr_subset = sjr_df[['Sourceid', 'SJR', 'SJR Best Quartile', 'Categories', 'Areas', 'H index']]
        sjr_subset = sjr_subset.drop_duplicates(subset=['Sourceid'], keep='first')
        # Merge with current results
        # Use 'left' join so it doesn't lose Scopus journals that aren't yet in SCImago
        target_df = self.filtered if self.filtered is not None else self.data_list_sources
        ranking_columns = ['SJR', 'SJR Best Quartile', 'Categories', 'Areas', 'H index']
        target_df = target_df.drop(
            columns=[column for column in ranking_columns if column in target_df.columns]
        )
        
        merged = target_df.merge(
            sjr_subset, 
            left_on='Sourcerecord ID', 
            right_on='Sourceid', 
            how='left'
        )
        
        # SCImago uses commas as decimal separators in CSV values.
        merged['SJR'] = pandas.to_numeric(
            merged['SJR'].astype('string').str.replace(',', '.', regex=False),
            errors='coerce',
        )
        # Cleanup: Remove the redundant 'Sourceid' column after merge
        if 'Sourceid' in merged.columns:
            merged = merged.drop(columns=['Sourceid'])

        self.filtered = merged
        print("Rankings added successfully.")

    def clean_up_columns(self):
        existing_cols = [c for c in self.columns_to_keep if c in self.filtered.columns]
        self.filtered = self.filtered[existing_cols]

    def print_filter_summary(self):
        print("="*10+" Summary of results "+"="*10)
        print("Number of rows:\t", len(self.filtered.index))
        print("="*10+"="*len(" Summary of results ")+"="*10)

    def export_filtered(self, filename):
        print(f"Exporting filtered list to excel file: {filename}")
        self.filtered.to_excel(filename, index=False)
    
    def save_with_autofit(self, filename):
        """Saves the filtered dataframe to Excel with auto-adjusted column widths."""
        if self.filtered is None:
            print("Nothing to save!")
            return
        
        # Use XlsxWriter as the engine
        with pandas.ExcelWriter(filename, engine='xlsxwriter') as writer:
            self.filtered.to_excel(writer, index=False, sheet_name='Filtered_Sources')
            
            workbook  = writer.book
            worksheet = writer.sheets['Filtered_Sources']

            # Iterate through columns and find the max length
            for i, col in enumerate(self.filtered.columns):
                # Find the maximum length of the column values + header
                column_len = self.filtered[col].astype(str).str.len().max()
                column_len = max(column_len, len(col)) + 2  # Add a little padding
                
                # Set the column width
                worksheet.set_column(i, i, column_len)
                
            # Add a simple filter to the top row
            worksheet.autofilter(0, 0, len(self.filtered), len(self.filtered.columns) - 1)

        print(f"File saved and formatted: {filename}")
    
    def save_with_formatting(self, filename):
        """Saves with Newlines, Auto-fit, and Q1-Q4 Color Coding."""
        if self.filtered is None: return

        # Add newlines to the Categories string
        if 'Categories' in self.filtered.columns:
            self.filtered['Categories'] = self.filtered['Categories'].astype(str).str.replace('; ', ';\n')
        
        # Use XlsxWriter to format the sheet
        with pandas.ExcelWriter(filename, engine='xlsxwriter') as writer:
            self.filtered.to_excel(writer, index=False, sheet_name='Rankings')
            
            workbook  = writer.book
            worksheet = writer.sheets['Rankings']

            # Define Formats
            wrap_top = workbook.add_format({'text_wrap': True, 'valign': 'top'})
            header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
            
            # Define Quartile Colors
            q1_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'}) # Green
            q2_format = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'}) # Yellow
            q3_format = workbook.add_format({'bg_color': '#FFCC99', 'font_color': '#9C0006'}) # Orange
            q4_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'}) # Red

            # Iterate through columns for Auto-fit and Wrapping
            for i, col in enumerate(self.filtered.columns):
                # Measure longest line in the column
                max_len = self.filtered[col].astype(str).apply(lambda x: max(len(line) for line in str(x).split('\n'))).max()
                if pandas.isna(max_len):
                    max_len = len(col) + 3
                else:
                    max_len = max(max_len, len(col)) + 3
                
                # Apply wrapping to everything for a clean look, set width
                worksheet.set_column(i, i, max_len, wrap_top)
                
                if col == "Source Title":
                    bold_format = workbook.add_format({'bold': True})
                    worksheet.conditional_format(1, i, len(self.filtered), i, {
                        'type':     'no_blanks',
                        'format':   bold_format
                    })

            # Apply Conditional Formatting to 'SJR Best Quartile' column
            if 'SJR Best Quartile' in self.filtered.columns:
                col_idx = self.filtered.columns.get_loc('SJR Best Quartile')
                # Apply to all rows except header
                row_range = f'1:{len(self.filtered)}' 
                
                worksheet.conditional_format(1, col_idx, len(self.filtered), col_idx, {
                    'type':     'cell',
                    'criteria': '==',
                    'value':    '"Q1"',
                    'format':   q1_format
                })
                worksheet.conditional_format(1, col_idx, len(self.filtered), col_idx, {
                    'type':     'cell',
                    'criteria': '==',
                    'value':    '"Q2"',
                    'format':   q2_format
                })
                worksheet.conditional_format(1, col_idx, len(self.filtered), col_idx, {
                    'type':     'cell',
                    'criteria': '==',
                    'value':    '"Q3"',
                    'format':   q3_format
                })
                worksheet.conditional_format(1, col_idx, len(self.filtered), col_idx, {
                    'type':     'cell',
                    'criteria': '==',
                    'value':    '"Q4"',
                    'format':   q4_format
                })

            # Freeze the top row so headers stay put
            # worksheet.freeze_panes(1, 0)

        print(f"Professional Excel file created: {filename}")