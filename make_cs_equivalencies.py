import pandas as pd
import os

csv_path = "data/ucl_degree_equivalencies_FINAL.csv"
output_path = "data/EquivalencyData.cs"

if not os.path.exists(csv_path):
    print(f"Error: Could not find {csv_path}")
    exit(1)

df = pd.read_csv(csv_path)
df = df[['country', 'Write 3.0', 'Write 2.2', 'Write 2.1', 'Write 1.0']].dropna(subset=['country'])

rows = []
for _, row in df.iterrows():
    c = str(row['country']).replace('"', '\\"').strip()
    g30 = str(row['Write 3.0']).replace('"', '\\"').strip()
    g22 = str(row['Write 2.2']).replace('"', '\\"').strip()
    g21 = str(row['Write 2.1']).replace('"', '\\"').strip()
    g10 = str(row['Write 1.0']).replace('"', '\\"').strip()
    rows.append(f'        ["{c}"] = new EquivalencyEntry {{ G30 = "{g30}", G22 = "{g22}", G21 = "{g21}", G10 = "{g10}" }},')

content = f"""using System.Collections.Generic;

namespace ADMerger.Services;

public class EquivalencyEntry
{{
    public string G30 {{ get; set; }} = "";
    public string G22 {{ get; set; }} = "";
    public string G21 {{ get; set; }} = "";
    public string G10 {{ get; set; }} = "";
}}

public static class EquivalencyData
{{
    public static readonly Dictionary<string, EquivalencyEntry> Equivalencies = new(System.StringComparer.OrdinalIgnoreCase)
    {{
{chr(10).join(rows)}
    }};
}}
"""

with open(output_path, "w") as out:
    out.write(content)
print(f"Successfully generated {output_path}")
