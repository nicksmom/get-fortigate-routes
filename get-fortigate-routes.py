import requests
import pandas as pd
import urllib3
import os

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def get_routes(fortigate_ip, api_token, vdom):
    url = f"https://{fortigate_ip}/api/v2/monitor/router/ipv4?vdom={vdom}"
    headers = {
        'Authorization': f'Bearer {api_token}'
    }

    try:
        response = requests.get(url, headers=headers, verify=False, timeout=5)
        response.raise_for_status()
        data = response.json()
        return data.get("results", [])
    except requests.Timeout:
        print(f"Warning: Failed to communicate with Fortigate at IP {fortigate_ip}. Request timed out.")
        return []
    except requests.RequestException as e:
        print(f"Failed to communicate with FortiGate at {fortigate_ip}. Error: {e}")
        return []

def highlight_routes(data_frame):
    # Define color map
    color_map = {
        'ospf': 'lightcoral',  # light red
        'bgp': 'lightblue',
        'connect': 'lightgreen',
        'static': 'plum'  # light purple
    }
    
    # Default coloring
    df_color = pd.DataFrame('background-color: white', index=data_frame.index, columns=data_frame.columns)
    
    # Apply coloring for entire row based on the 'type' column
    for route_type, color in color_map.items():
        df_color[data_frame['type'] == route_type] = f'background-color: {color}'
    
    return df_color


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

    output_file = 'fortigate_routes.xlsx'
    
    if not os.path.exists(output_file):
        try:
            open(output_file, 'w').close()
        except Exception as e:
            print(f"Couldn't create the file due to: {e}")

    with pd.ExcelWriter('fortigate_routes.xlsx') as writer:
        for index, row in fortigates.iterrows():
            for vdom in vdoms:
                routes = get_routes(row['ip'], row['token'], vdom)
                if routes:
                    df = pd.DataFrame(routes)
                    
                    # Apply styling to the DataFrame
                    styled_df = df.style.apply(highlight_routes, axis=None)
                    
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
