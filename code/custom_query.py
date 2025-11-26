from pandas_gbq import read_gbq

def read_sql_query(file_path):
    try:
        with open(file_path, 'r') as file:
            return file.read()
    except (FileNotFoundError, IOError) as e:
        print(f"Error reading '{file_path}': {e}")
        return None

# File paths to SQL query files
sql_files = {
    "weekly_note": '../query/weekly_query.txt',
    "finance_data": '../query/Finance_Query.txt',
    "deal_share": '../query/deal_share_query.txt',
    "gds_incentives": '../query/gds_incentives_query.txt',
    "tsa_data": '../query/tsa_data_query.txt',
    "direct_parity_data": '../query/direct_parity_data_query.txt',
    "meta_parity_data_query": '../query/meta_parity_data_query.txt',
    "upsell_query": '../query/upsell_query.txt',
    "bookability_query": '../query/bookability_query.txt',
    "roi_query": '../query/roi_query.txt',
    "dau_conversion_query": '../query/dau_conversion_query.txt',
    "midt_query": '../query/midt_data.txt',
    "sem_query": '../query/sem_query.txt'                 
}