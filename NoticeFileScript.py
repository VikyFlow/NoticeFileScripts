import pandas as pd
# Upload Excel file
file_path = "sbom.app.xlsx"  # 
try:
    df = pd.read_excel(file_path, engine='openpyxl')
except FileNotFoundError:
    print("Errore: Il file specificato non è stato trovato. Verifica il percorso del file.")
    exit()
except Exception as e:
    print(f"Errore durante il caricamento del file Excel: {e}")
    exit()
 
# Checks for mandatory columns and prints an error message if they are missing
required_columns = ["Author", "Name", "Version", "License URL", "Copyright"]
missing_columns = [column for column in required_columns if column not in df.columns]
if missing_columns:
    print(f"Errore: Le seguenti colonne sono mancanti nel file Excel: {', '.join(missing_columns)}")
    exit()
 
# Generate the “Component List” in HTML
html_content = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Third-Party Software Disclosure</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            font-size: 14px;
            text-align: left;
        }
        table, th, td {
            border: 1px solid #2F4F4F;
            padding: 10px;
        }
        th {
            background-color: #f2f2f2;
        }
        tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        a {
            color: #007BFF;
            text-decoration: none;
        }
        a:hover {
            text-decoration: underline;
        }
    </style>
</head>
<body>
    <h2>Third-Party Software Disclosure Document</h2>
    <p>This document provides information on third-party components used in this product, along with their respective licenses and authors where available.</p>
    <table>
        <thead>
            <tr>
                <th>Component Name</th>
                <th>Version</th>
                <th>License Name</th>
                <th>Author</th>
            </tr>
        </thead>
        <tbody>
"""
 
# Iterate on the DataFrame rows and add the HTML rows.
for _, row in df.iterrows():
    component_name = row.get("Name", "N/A")
    version = row.get("Version", "N/A")
    license_name = row.get("License Name", "N/A")
    license_url = row.get("License URL", "N/A")
    author = row.get("Author", "")
 
    # If the value “Author” is empty or NaN, show it as “-”
    author_text = author if pd.notna(author) and author != "" else "-"

 
    # Add the row to the HTML table
    html_content += f"""
            <tr>
                <td>{component_name}</td>
                <td>{version}</td>
                <td><a href="{license_url}" target="_blank">{license_name}</a></td>
                <td>{author_text}</td>
            </tr>
    """

# Close HTML content
html_content += """
        </tbody>
    </table>
</body>
</html>
"""

 
# Save the HTML file
output_file = "component_list_right.html"
try:
    with open(output_file, "w", encoding="utf-8") as file:
        file.write(html_content)
    print(f"La 'Component List' è stata generata e salvata in {output_file}.")
except Exception as e:
    print(f"Errore durante il salvataggio del file: {e}")
