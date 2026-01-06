# © Ricki Angel 2026 | TechAngelX
# Converts institution alias mappings CSV into a C# dictionary for lookup.


import csv
import os
import sys

mapping_path = "data/institution_mappings.csv"
output_path = "data/MappingData.cs"

if not os.path.exists(mapping_path):
    print(f"Error: Could not find {mapping_path}")
    sys.exit(1)

# Try common encodings to avoid the 'utf-8' decode error
encodings = ['utf-8-sig', 'latin-1', 'cp1252']
content_rows = None

for enc in encodings:
    try:
        with open(mapping_path, mode='r', encoding=enc) as f:
            reader = csv.reader(f)
            # Try to read all rows to verify encoding
            content_rows = list(reader)
            print(f"Successfully read file using {enc} encoding.")
            break
    except (UnicodeDecodeError, LookupError):
        continue

if content_rows is None:
    print("Error: Could not decode the CSV file with UTF-8, Latin-1, or CP1252.")
    sys.exit(1)

try:
    # Skip header row if it exists
    if content_rows and ("alias" in content_rows[0][0].lower() or "institution" in content_rows[0][0].lower()):
        content_rows = content_rows[1:]
            
    rows = []
    for row in content_rows:
        if len(row) >= 2:
            # Clean up smart quotes/special chars and escape for C#
            alias = row[0].strip().replace('"', '\\"')
            full_name = row[1].strip().replace('"', '\\"')
            if alias and full_name:
                rows.append(f'        ["{alias}"] = "{full_name}",')

    content = f"""using System.Collections.Generic;

namespace ADMerger.Services;

public static class MappingData
{{
    public static readonly Dictionary<string, string> InstitutionMappings = new(System.StringComparer.OrdinalIgnoreCase)
    {{
{chr(10).join(rows)}
    }};
}}
"""
    with open(output_path, "w") as out:
        out.write(content)
    print(f"Successfully generated {output_path} with {len(rows)} mappings.")

except Exception as e:
    print(f"Error: {e}")
