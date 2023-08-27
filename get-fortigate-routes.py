import requests
import pandas as pd
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def get_routes(ip, token, vdom):
    headers = {
        'Authorization': f'Bearer {token}'
    }
    url = f'https://{ip}/api/v2/monitor/router/ipv4?vdom={vdom}'
    try:
        response = requests.get(url, headers=headers, verify=False, timeout=5)  # timeout set to 5 seconds
        response.raise_for_status()
        return response.json()['results']
    except requests.Timeout:
        print(f"Warning: Failed to communicate with Fortigate at IP {ip}. Request timed out.")
        return []
    except requests.RequestException as e:
        print(f"Error communicating with Fortigate at IP {ip}. Error: {e}")
        return []

def highlight_routes(val):
    color_map = {
        'ospf': 'lightcoral',  # light red
        'bgp': 'lightblue',
        'connected': 'lightgreen',
        'static': 'plum'  # light purple
    }
    color = color_map.get(val, 'white')  # default to white if route type not in color_map
    return f'background-color: {color}'

def main():
    # Prompt user for VDOMs
    vdoms_input = input("Enter comma separated list of VDOMs (default is 'root'): ")
    vdoms = [vdom.strip() for vdom in vdoms_input.split(",")] if vdoms_input else ["root"]

    # Prompt user for CSV filename
    csv_filename = input("Enter the name of the CSV file (e.g., fortigates.csv): ")

    # Load data from CSV
    try:
        fortigates = pd.read_csv(csv_filename, delim_whitespace=True)
    except Exception as e:
        print(f"Error reading CSV: {e}")
        return

    if 'ip' not in fortigates.columns or 'token' not in fortigates.columns:
        print("The CSV file is missing required columns ('ip' and/or 'token').")
        return

    with pd.ExcelWriter('fortigate_routes.xlsx') as writer:
        for index, row in fortigates.iterrows():
            for vdom in vdoms:
                routes = get_routes(row['ip'], row['token'], vdom)
                if routes:
                    df = pd.DataFrame(routes)
                    
                    # Apply styling to the DataFrame
                    styled_df = df.style.applymap(highlight_routes, subset=['type'])
                    
                    # Write styled DataFrame to Excel
                    sheet_name = f"{row['name']}_{vdom}"
                    styled_df.to_excel(writer, sheet_name=sheet_name, index=False, engine='openpyxl')
                    
                    # Get the sheet to apply additional formatting
                    worksheet = writer.sheets[sheet_name]
                    
                    # Enable autofilter
                    worksheet.auto_filter.ref = worksheet.dimensions
                    
                    # Autofit columns based on column content
                    for column in worksheet.columns:
                        max_length = 0
                        column = [cell for cell in column]
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(cell.value)
                            except:
                                pass
                        adjusted_width = (max_length + 2)
                        worksheet.column_dimensions[column[0].column_letter].width = adjusted_width


if __name__ == "__main__":
    main()
