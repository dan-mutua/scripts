import zipfile
import os
import pandas as pd
from bs4 import BeautifulSoup

def extract_saz_file(saz_file_path, extract_to_folder):
    zip_file_path = saz_file_path.replace('.saz', '.zip')
    os.rename(saz_file_path, zip_file_path)
    
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to_folder)
    
    # Cleanup
    os.rename(zip_file_path, saz_file_path)

def parse_index_file(index_file_path):
    with open(index_file_path, 'r', encoding='utf-8') as file:
        html_content = file.read()

    soup = BeautifulSoup(html_content, 'html.parser')

    table = soup.find('table')

    results = []

    for row in table.find('tbody').find_all('tr'):
        cols = row.find_all('td')

        process = cols[9].text.strip()

        if any(proc in process for proc in ['powerpnt', 'winword', 'excel']):
            flags = cols[12].text.strip()

            x_hostip = None
            https_client_snihostname = None

            for flag in flags.split(';'):
                if 'x-hostip' in flag:
                    x_hostip = flag.split(':')[1].strip()
                if 'https-client-snihostname' in flag:
                    https_client_snihostname = flag.split(':')[1].strip()

            row_data = {
                'Process': process,
                'x-hostip': x_hostip,  
                'https-client-snihostname': https_client_snihostname,
            }
            results.append(row_data)

    return results

def main(saz_file_path):
    extract_to_folder = os.path.join(os.path.dirname(__file__), '..', 'data', 'extracted_files')
    
    output_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'data', 'output'))
    
    os.makedirs(extract_to_folder, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    
    extract_saz_file(saz_file_path, extract_to_folder)
    
    index_file_path = os.path.join(extract_to_folder, '_index.htm')
    results = parse_index_file(index_file_path)
    
    df = pd.DataFrame(results)
    
    df.columns = df.columns.str.strip()
    
    required_columns = ['Process', 'x-hostip', 'https-client-snihostname', 'arinResult']
    for col in required_columns:
        if col not in df.columns:
            df[col] = None
    
   # print("DataFrame columns:", df.columns.tolist())
   # print("\nFirst few rows of data:")
   # print(df.head())
    
    output_path = os.path.join(output_dir, 'filtered_results.xlsx')
    
    
    df.to_excel(output_path, index=False)
    print(f"\nFile saved at: {output_path}")
    
    scripts_copy = os.path.join(os.path.dirname(__file__), 'filtered_results.xlsx')
    df.to_excel(scripts_copy, index=False)
    print(f"Copy saved at: {scripts_copy}")
    
    return output_path

if __name__ == '__main__':
    import sys
    if len(sys.argv) > 1:
        saz_file_path = sys.argv[1]
        main(saz_file_path)
    else:
        print("Error: Please provide a .saz file path")
