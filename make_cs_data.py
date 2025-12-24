import os
import glob
import sys

# Search for any ranking file in the data folder
search_pattern = os.path.join("data", "*THE Ranking 2026*")
matching_files = glob.glob(search_pattern)

if not matching_files:
    print("Error: Could not find ranking file in 'data' folder.")
    sys.exit(1)

file_path = matching_files[0]
print(f"Reading file: {file_path}")

try:
    if file_path.endswith('.xlsx'):
        import pandas as pd
        df = pd.read_excel(file_path)
    else:
        import pandas as pd
        df = pd.read_csv(file_path)

    rows = []
    # Using iloc to be index-based (0=Rank, 1=Institution)
    for _, row in df.iterrows():
        name = str(row.iloc[1]).replace('"', '\\"').strip()
        rank = str(row.iloc[0]).strip()
        rows.append(f'        ["{name}"] = "{rank}",')

    content = f"""using System.Collections.Generic;

namespace ADMerger.Services;

public static class RankingData
{{
    public static readonly Dictionary<string, string> Rankings = new(System.StringComparer.OrdinalIgnoreCase)
    {{
{chr(10).join(rows)}
    }};
}}
"""
    with open("data/RankingData.cs", "w") as out:
        out.write(content)
    print(f"Successfully generated Services/RankingData.cs with {len(rows)} entries.")

except ImportError:
    print("Error: 'pandas' and 'openpyxl' are required. Run: pip3 install pandas openpyxl")
except Exception as e:
    print(f"Error: {e}")
