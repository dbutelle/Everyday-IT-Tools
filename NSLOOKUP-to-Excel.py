#### DODOW62 ####
#### Version 1.0 ####


import subprocess
import pandas as pd
import os

def nslookup(entry):
    """Function to perform nslookup on an entry (domain or IP) and return the output."""
    try:
        result = subprocess.run(['nslookup', entry], capture_output=True, text=True, check=True)
        return result.stdout
    except subprocess.CalledProcessError as e:
        return str(e)

def parse_nslookup_output(output):
    """Function to parse nslookup output and extract relevant information."""
    lines = output.split('\n')
    parsed_data = {'Entry': '', 'Server': '', 'Server Address': '', 'Name': '', 'Address': ''}

    for line in lines:
        if 'Server:' in line or 'Serveur :' in line:
            parsed_data['Server'] = line.split(': ')[1]
        elif 'Address:' in line and not parsed_data['Server Address']:
            parsed_data['Server Address'] = line.split(': ')[1]
        elif 'Name:' in line or 'Nom :' in line:
            parsed_data['Name'] = line.split(': ')[1]
        elif 'Address:' in line and parsed_data['Server Address']:
            parsed_data['Address'] = line.split(': ')[1]

    return parsed_data

def nslookup_entries_to_excel(entries, output_file):
    """Function to perform nslookup on a list of entries and export results to an Excel file."""
    results = []

    for entry in entries:
        output = nslookup(entry)
        parsed_data = parse_nslookup_output(output)
        parsed_data['Entry'] = entry  # Ensure the entry is correctly set
        results.append(parsed_data)

    # Create a DataFrame and export to Excel
    df = pd.DataFrame(results)
    df.to_excel(output_file, index=False)
    print(f"NSLookup results have been exported to {output_file}")

# Main function to take user input and perform nslookup
def main():
    current_dir = os.path.dirname(__file__)
    entries = []
    print("Enter IP addresses or domain names (enter 'quit' or 'exit' to finish):")
    while True:
        entry = input("IP address or domain name: ")
        if entry.lower() == 'quit' or entry.lower() == 'exit':
            break
        entries.append(entry)
    
    if entries:
        output_file = os.path.join(current_dir, "nslookup_results.xlsx")
        nslookup_entries_to_excel(entries, output_file)
    else:
        print("No IP addresses or domain names were entered.")

if __name__ == "__main__":
    main()
