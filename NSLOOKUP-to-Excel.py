import subprocess
import pandas as pd
import os

def nslookup(entry):
    """Function to perform nslookup on an entry (domain or IP) and return the output."""
    try:
        result = subprocess.run(['nslookup', entry], capture_output=True, text=True, check=True)
        print(result)
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
    print(f"Les résultats de NSLookup ont été exportés vers {output_file}")

# Main function to take user input and perform nslookup
def main():
    current_dir = os.path.dirname(__file__)
    entries = []
    print("Entrez les adresses IP ou les noms de domaine (entrez 'quit or exit' pour terminer) :")
    while True:
        entry = input("Adresse IP ou nom de domaine : ")
        if entry.lower() == 'quit' or entry.lower() == 'exit' :
            break
        entries.append(entry)
    
    if entries:
        output_file = os.path.join(current_dir, "nslookup_results.xlsx")
        nslookup_entries_to_excel(entries, output_file)
    else:
        print("Aucune adresse IP ou nom de domaine n'a été saisie.")

if __name__ == "__main__":
    main()
