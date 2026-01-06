# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build and Test Commands

### Running the Application
```bash
dotnet run
```

### Running Tests
```bash
# Run all tests
dotnet test

# Run a single test class
dotnet test --filter "FullyQualifiedName~RankingServiceTests"

# Run a specific test method
dotnet test --filter "FullyQualifiedName~RankingServiceTests.GetRanking_ExactMatch_ReturnsRank"
```

### Building for Distribution

**Windows (PowerShell):**
```powershell
.\goDeploy.ps1
```
Creates a standalone executable at `Desktop\ADMerger.exe` with all data files embedded. The script:
- Targets .NET 10.0 and win-x64 runtime
- Closes running ADMerger and Excel processes
- Publishes as a single self-contained file
- Embeds all ranking and mapping data

**macOS (Bash):**
```bash
./ADMACBUILDER.sh
```
Creates `ADMerger.app` bundle on Desktop with osx-x64 runtime.

### Cleaning Build Artifacts
```bash
dotnet clean
```

## Architecture Overview

### Core Services Architecture

The application follows a service-oriented architecture with dependency injection. Services are instantiated in `MainWindow.axaml.cs:115-123` without a formal DI container:

```csharp
_csvService = new CsvService();
_equivalencyService = new EquivalencyService();
_matchingService = new InstitutionMatchingService();
_rankingService = new RankingService(_matchingService);
_gradeService = new GradeClassificationService(_equivalencyService);
```

**Key Service Dependencies:**
- `RankingService` depends on `IInstitutionMatchingService`
- `GradeClassificationService` depends on `IEquivalencyService`
- All other services are independent

### Data Flow Pipeline

1. **Input Loading** (`CsvService`):
   - `LoadInTrayRecords()` - Student numbers and received dates
   - `LoadApplicationRecords()` - Detailed applicant information

2. **Data Enrichment** (MainWindow.axaml.cs:232-270):
   - Cross-reference by Student ID
   - Lookup university ranking via `RankingService.GetRanking()`
   - Determine UK grade classification via `GradeClassificationService.DetermineUKClassification()`
   - Format dates using `DateFormatter`
   - Map programme codes using `ProgrammeMapping`

3. **Output Generation** (`CsvService.GenerateOutputFiles()`):
   - Creates programme-specific CSV files
   - Includes enriched ranking and grade data

### Ranking Lookup Algorithm

The ranking lookup follows a three-stage waterfall (RankingService.cs:28-46):

1. **Institution Mapping** - Apply predefined mappings from `MappingData.InstitutionMappings`
   - Example: "UCL" → "University College London"
   - Example: "MIT" → "Massachusetts Institute of Technology"

2. **Exact Match** - Case-insensitive dictionary lookup in `RankingData.Rankings`

3. **Fuzzy Match** - Uses `InstitutionMatchingService.FindBestMatch()` with:
   - Levenshtein distance algorithm
   - Word-level matching (ignoring common words like "university", "college")
   - 70% minimum similarity threshold (InstitutionMatchingService.cs:12)
   - Exact word matches score 100%, fuzzy matches score 50%

Returns "NR" (Not Ranked) if no match is found.

### Data Storage Pattern

All reference data is **hardcoded in C# classes** in the `data/` folder:
- `RankingData.cs` - THE World University Rankings (2026) as a dictionary
- `MappingData.cs` - Institution name mappings (abbreviations, joint degrees)
- `EquivalencyData.cs` - Country-specific grade equivalencies

**Important:** Unlike the README's description, there are no external CSV or Excel files loaded at runtime. All data is compiled into the application. This is why deployment creates a single standalone executable.

### UI Framework

The application uses **Avalonia UI 11.2.2** (not Windows Forms as README suggests):
- Cross-platform XAML-based UI
- MainWindow in `Views/MainWindow.axaml.cs`
- App initialization in `App.axaml.cs`
- Observable collections for live data binding
- Custom `ProcessingItem` class implements `INotifyPropertyChanged` for real-time UI updates

### Testing Strategy

Tests use **xUnit** with **Moq** for mocking:
- Service unit tests focus on business logic isolation
- Mocking pattern: `Mock<IInstitutionMatchingService>` for testing `RankingService`
- Tests verify null handling, exact matches, mappings, and fuzzy matching
- No integration tests with external files (data is hardcoded)

## Key Implementation Details

### Fuzzy Matching Thresholds
- Minimum match threshold: 70% (InstitutionMatchingService.cs:12)
- Exact word match: 100% score
- Fuzzy word match: 50% score (weighted at half)
- Words ≤ 2 characters are ignored
- Common words filtered: "university", "college", "institute", "school", etc.

### Date Handling
- Uses `DateFormatter.FormatDate()` for input parsing
- Uses `DateFormatter.CalculateDueDate()` for deadline computation
- Due date logic is in the Utilities layer

### Programme Code Mapping
- `ProgrammeMapping.GetCode()` converts full programme names to codes
- Used during output record generation (MainWindow.axaml.cs:248)

### Audio Feedback
- Plays confirmation sound on completion (MainWindow.axaml.cs:349-386)
- Platform-specific: `afplay` on macOS, `mciSendString` on Windows
- Audio file: `audio/confirmed.mp3` (copied to output during build)

### Processing Flow Control
- Cancellable processing using `CancellationTokenSource`
- UI updates via `Dispatcher.UIThread.InvokeAsync()` for thread safety
- Progress bar and status updates in real-time
- Summary modal shows grade distribution on completion

## Important Caveats

1. **Data is Hardcoded**: Despite README references to CSV/Excel files, all ranking and mapping data is compiled into the application as C# dictionaries in the `data/` folder.

2. **No DI Container**: Services are manually instantiated in MainWindow constructor. To add a new service, update both the instantiation chain and any dependent services.

3. **Fuzzy Matching Limitations**: The README mentions potential issues with similar Chinese university names (finance/economics institutions). Consider adding explicit mappings to `MappingData.cs` for problematic cases.

4. **Platform-Specific Audio**: The sound playback code has platform-specific implementations and may fail silently on unsupported platforms.

5. **Version Numbering**: Version uses build year: `v1.0.{YY}` (e.g., v1.0.25 for 2025). See ADMerger.csproj:13-16 and MainWindow.axaml.cs:310-318.

## Testing Notes

When writing tests:
- Mock `IInstitutionMatchingService` when testing `RankingService`
- Mock `IEquivalencyService` when testing `GradeClassificationService`
- Reference known universities from `RankingData.cs` (Oxford, Cambridge, MIT, etc.)
- Test abbreviation mappings like "UCL", "MIT" which are in `MappingData`
- Verify "NR" return value for unknown/unmapped institutions
