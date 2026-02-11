# ADMerger

![ADMerger Interface](readme_images/screenshot1.png)

![Optional Fields](readme_images/screenshot2.png)
## Overview

ADMerger accesses two internal JSON files containing student application data and automatically enriches them with:
- THE World University Rankings (2026)
- UK degree classification equivalencies
- Country-specific grade conversions
- Programme codes and mappings

The application uses intelligent fuzzy matching with Levenshtein distance algorithm to handle institution name variations and abbreviations.

## Features

- **Automated Ranking Lookup**: Matches institution names against THE World University Rankings using exact and fuzzy matching
- **Configurable Mappings**: Institution name variations, abbreviations, and joint degrees managed via CSV configuration
- **Grade Classification**: Automatic UK grade classification based on country-specific equivalency rules
- **Encoding Normalization**: Handles various text encodings and special characters in international names
- **Modern UI**: Clean Windows Forms interface with drag-and-drop file loading
- **Comprehensive Logging**: Debug logs track all matching decisions for troubleshooting

## Architecture

### Project Structure

```
ADMerger/
+-- Program.cs                      # Application entry point with DI setup
+-- MainForm.cs                     # Main UI form
+-- Models/
�   +-- InTrayRecord.cs            # Document 1 data model
�   +-- ApplicationRecord.cs        # Document 2 data model
�   +-- OutputRecord.cs            # Merged output data model
�   +-- DegreeEquivalency.cs       # Country equivalency model
+-- Services/
�   +-- CsvService.cs              # CSV reading/writing with CsvHelper
�   +-- EquivalencyService.cs      # Country grade equivalencies
�   +-- RankingService.cs          # THE Rankings lookup with caching
�   +-- GradeClassificationService.cs  # UK grade determination
�   +-- InstitutionMatchingService.cs  # Fuzzy matching engine
+-- UI/Controls/
�   +-- ModernButton.cs            # Styled button component
�   +-- ModernFilePanel.cs         # Drag-drop file selector
+-- Utilities/
�   +-- DateFormatter.cs           # Date parsing and formatting
�   +-- TextNormalizer.cs          # Text cleaning and normalization
+-- Configuration/
�   +-- ProgrammeMapping.cs        # Programme code mappings
+-- data/
    +-- THE Ranking 2026.xlsx      # University rankings data
    +-- ucl_degree_equivalencies_FINAL.csv  # Country equivalencies
    +-- institution_mappings.csv   # Institution name mappings
```

### Design Principles

- **SOLID Architecture**: Interface-based design with dependency injection
- **Open-Closed Principle**: Institution mappings externalized to CSV for extension without code modification
- **Separation of Concerns**: Clear separation between UI, business logic, and data access
- **Service-Oriented**: All core functionality encapsulated in injected services

## Fuzzy Matching Algorithm

The application uses a sophisticated Levenshtein distance-based fuzzy matching system:

1. **Normalization**: Removes diacritics, common words (university, college), and special characters
2. **Word-Level Matching**: Compares significant words between search and candidate names
3. **Exact Priority**: Exact word matches score 100%, fuzzy matches score 50%
4. **Threshold Filtering**: Only matches scoring =75% are considered
5. **Special Handling**: NOT_RANKED entries bypass fuzzy matching to prevent false positives

### Matching Examples

```
"University of Warwick" ? "University of Warwick" (exact match, 100%)
"LSE" ? "London School of Economics and Political Science" (mapping)
"St. Andrews" ? "St Andrews" (fuzzy match, 95%)
```

## Configuration

### Institution Mappings (`data/institution_mappings.csv`)

Handles name variations without code changes:

```csv
SourceName,TargetName,MappingType
UCL,University College London,Abbreviation
HKUST,The Hong Kong University of Science and Technology,Abbreviation
Xi'an Jiaotong-Liverpool University,University of Liverpool,JointDegree
Southwestern University of Finance and Economics,NOT_RANKED,NotRanked
```

**Mapping Types:**
- **Abbreviation**: Common abbreviations (MIT, UCLA, IIT Madras)
- **JointDegree**: Transnational programs map to UK partner ranking
- **NotRanked**: Institutions not in THE Rankings (prevents false matches)
- **AltSpelling**: Regional spelling variations

### Ranking Status Values

- **Numeric**: Single rank (1, 52, 143)
- **Range**: Band ranking (601-800, 1001-1200)
- **Reporter**: Participated in THE process but not ranked
- **NR**: Not found in THE Rankings

## Building and Distribution

### Development Build

```bash
dotnet run
```

### Production Build (Standalone EXE)

```powershell
.\goDeploy.ps1
```

Creates a single-file executable at `Desktop\ADMerger.exe` with all data files embedded.

**Build Configuration:**
- Target: .NET 9.0 Windows
- Self-contained: Yes
- Single file: Yes (with native libraries extracted)
- Embedded resources: All data files (Excel, CSVs)

## Usage

![ADMerger Application Interface](readme_images/screenshot1.png)

1. **Load Document 1**: SELECT In-tray file (latest applicants CSV)
2. **Load Document 2**: Department Application reports CSV (detailed applicant data)
3. **Process Files**: Cross-references by Student No., enriches with rankings and grades
4. **Review Output**: Programme-specific CSV files with all enriched data

### Input Requirements

**Document 1** must contain:
- Student No.
- Received on date

**Document 2** must contain:
- Applicant ID (matches Student No.)
- Programme, Name, 
- Fee Status, Country of Study
- Qualification, Subject, Institution, Grade
- Equivalency notes

### Output Format

Generates one CSV per programme with columns:
- ReceivedDate, DueDate, StudentNo, Programme
- Forename, Surname,  
- FeeStatus, 
- QualificationName, DegreeSubject, InstitutionName
- **THERanking** (enriched)
- CountryOfStudy, EquivalencyNote, OverallGradeGPA
- **UKGrade** (enriched)

## Dependencies

- **CsvHelper** (33.1.0): CSV parsing with class mapping
- **EPPlus** (7.5.2): Excel file reading for THE Rankings

## Logging

Debug logs written to `Desktop\ranking_matches.log`:

```
2025-11-24 20:34:06 - Original: 'University of Oxford' | Normalized: 'University of Oxford' | InDict: True
2025-11-24 20:34:06 - EXACT MATCH: 'University of Oxford' = Rank 1
2025-11-24 20:34:06 - FUZZY MATCH: 'St. Andrews' -> 'St Andrews' = Rank 162
2025-11-24 20:34:06 - NOT RANKED: 'Juniata College' (marked as not in THE Rankings)
```

## Maintenance

### Adding New Institution Mappings

1. Edit `data/institution_mappings.csv`
2. Add line: `SourceName,TargetName,MappingType`
3. Restart application (no rebuild required)

### Updating Rankings

1. Replace `data/THE Ranking 2026.xlsx`
2. Ensure Column A = Rank, Column B = Institution Name
3. Restart application

### Adjusting Fuzzy Match Threshold

Edit `Services/InstitutionMatchingService.cs`:
```csharp
private const int MinimumMatchThreshold = 75; // Increase for stricter matching
```

## Known Limitations

- Fuzzy matching may incorrectly match similar Chinese university names (e.g., finance/economics institutions)
- Joint degree identification relies on specific keywords (e.g., "-Liverpool" in institution name)
- THE Rankings "Reporter" status requires manual verification



## License
Creative Commons:
© Ricki Angel 2026 | TechAngelX

This work is licensed under a Creative Commons Attribution-NonCommercial-NoDerivatives 4.0 International License.

## Disclaimer

This project is for personal or educational use only and comes without any warranty.
***
<h2 align="center">Support</h2>
<div align="center">
  For issues or questions, feel free to reach out to me on GitHub.
  <br /><br />
  <a href="https://techangelx.com" target="_blank" rel="noopener noreferrer">
    <img src="./readme_images/logo.png" alt="Tech Angel X Logo" width="70" height="70" style="border-radius: 50%; border: 4px solid #ffffff; box-shadow: 0 4px 10px rgba(0,0,0,0.2);">
  </a>
  <br /><br />
  <b>Built by Ricki Angel</b> • <a href="https://techangelx.com" target="_blank" rel="noopener noreferrer">Tech Angel X</a>
</div>




