import pandas as pd
import re

def extract_emails(excel_path, txt_file_path):
    try:
        # Load CAPIDs from Excel file
        df = pd.read_excel(excel_path)
        capids = set(df['CAPID'].astype(str))

        # Adjusted regular expression to match only primary emails
        email_pattern = re.compile(r'"(\d+)","EMAIL","PRIMARY","([^"]+)"')
        emails = []

        with open(txt_file_path, 'r') as file:
            for line in file:
                match = email_pattern.search(line)
                if match:
                    capid, email = match.groups()
                    if capid in capids:
                        emails.append(email)

        # Output only the primary emails
        for email in emails:
            print(email)

    except Exception as e:
        print("An error occurred:", e)

# Example usage
extract_emails(
    'path_to_excel_file.xlsx',  # Path to the Excel file with the CAPID's in them
    'path_to_text_file.txt'     # Path to the text file
)
