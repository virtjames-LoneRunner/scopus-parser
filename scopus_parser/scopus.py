import requests
import pandas
import re

from scopus_parser.config import BISUScopusConfig


class BISUScopus():
    list_link = ""
    list_filename = ""
    list_file = None
    list_file_sheet_names = []
    
    data_list_sources = None
    data_asjc = None
    filtered = None

    def __init__(self, config: BISUScopusConfig):
        self.list_link = config.list_link
        self.list_filename = config.list_filename
        return
    
    def retrieve_list(self):
        if (self.list_link == ""):
            print("Please provide the link")
            return False
        response = requests.get(self.list_link) 

        if response.status_code == 200:
            with open(self.list_filename, 'wb') as f:
                f.write(response.content)
                print(f"Successfully written to {self.list_filename}.")
                return True
        else:
            print("There was an error while downloading the scopus list file.")

        return False

    def load_file(self):
        print("Loading list file....")
        self.list_file = pandas.ExcelFile(self.list_filename)
        self.list_file_sheet_names = self.list_file.sheet_names
    
    def read_sources(self):
        print("Reading scopus sources from list....")
        self.data_list_sources = pandas.read_excel(self.list_filename)

        df = pandas.read_excel(self.list_file, self.list_file_sheet_names[-1])
        mapping_df = df.dropna(subset=df.columns[:2], how='any')

        # Clean up: Ensure the 'Code' column is numeric and Descriptions are strings
        # This ignores header rows like "Code" or "Description"
        mapping_df.columns = ['Code', 'Description'] + list(mapping_df.columns[2:])
        mapping_df['Code'] = pandas.to_numeric(mapping_df['Code'], errors='coerce')
        mapping_df = mapping_df.dropna(subset=['Code'])

        # Final Clean Table (Only first two columns)
        asjc_lookup_table = mapping_df[['Code', 'Description']].reset_index(drop=True)
        self.data_asjc = dict(zip(asjc_lookup_table['Code'], asjc_lookup_table['Description']))

    def filter_by_column(self, column, keywords = []):
        # Add \b to ensure we only match whole words
        pattern = '|'.join([fr'\b{re.escape(word)}\b' for word in keywords])

        if self.filtered is None:
            self.filtered = self.data_list_sources[self.data_list_sources[column].str.contains(pattern, case=False, na=False)]
        else:
            self.filtered = self.filtered[self.filtered[column].str.contains(pattern, case=False, na=False)]
        
        #return self.filtered
    
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
        
        sjr_df = pandas.read_csv(scimago_csv_path, sep=';')

        sjr_subset = sjr_df[['Sourceid', 'SJR', 'SJR Best Quartile', 'Categories', 'Areas', 'H index']]
        # Merge with current results
        # Use 'left' join so it doesn't lose Scopus journals that aren't yet in SCImago
        target_df = self.filtered if self.filtered is not None else self.data_list_sources
        
        merged = target_df.merge(
            sjr_subset, 
            left_on='Sourcerecord ID', 
            right_on='Sourceid', 
            how='left'
        )
        
        # 4. Cleanup: Remove the redundant 'Sourceid' column after merge
        if 'Sourceid' in merged.columns:
            merged = merged.drop(columns=['Sourceid'])

        self.filtered = merged
        print("Rankings added successfully.")

    def clean_up_columns(self):
        columns_to_keep = [
            'Sourcerecord ID',
            'SJR', 
            'SJR Best Quartile', 
            'Source Title', 
            'Categories', 
            'Areas', 
            'H index'
            'All Science Journal Classification Codes (ASJC)',
            'Field Descriptions',
            'ISSN', 
            'EISSN', 
            'Active or Inactive',
            'Coverage',
            'Title Discontinued by Scopus',
            'Article Language in Source (Three-Letter ISO Language Codes)',
            'Open Access Status',
            'Publisher',
        ]
        existing_cols = [c for c in columns_to_keep if c in self.filtered.columns]
        
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
                
            # Optional: Add a simple filter to the top row
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
                max_len = max(max_len, len(col)) + 3
                
                # Apply wrapping to everything for a clean look, set width
                worksheet.set_column(i, i, max_len, wrap_top)

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