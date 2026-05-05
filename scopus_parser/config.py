import os

class BISUScopusConfig():
    list_link="https://downloads.ctfassets.net/o78em1y1w4i4/7xtaTxNiNcWRTeZkV86eNy/de9e757c475827b03206a5bf4d24c8a3/ext_list_Jan_2026.xlsx"
    list_filename = os.path.join("..", "scopus", "scopus_list.xlsx")
    columns_to_keep = [
        'Sourcerecord ID',
        'SJR', 
        'SJR Best Quartile', 
        'Source Title', 
        'Categories', 
        'Areas', 
        'Source Type',
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