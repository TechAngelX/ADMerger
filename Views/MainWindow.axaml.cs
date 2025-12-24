// Views/MainWindow.axaml.cs

using Avalonia; using Avalonia.Controls; using Avalonia.Interactivity; using Avalonia.Platform.Storage; using Avalonia.Threading; using Avalonia.Media; using System.Collections.ObjectModel; using ADMerger.Services; using ADMerger.Models; using ADMerger.Configuration;

using ADMerger.Utilities; using System; using System.Collections.Generic; using System.IO; using System.Linq; using System.Reflection; using System.Threading; using System.Threading.Tasks; using System.Runtime.InteropServices; using System.Diagnostics; using System.ComponentModel; using System.Text;

namespace ADMerger.Views;

public class ProcessingItem : INotifyPropertyChanged { private string _status = ""; private string _ukGrade = ""; private string _studentNo = ""; private string _name = "";

public string StudentNo { get => _studentNo; set { _studentNo = value; OnPropertyChanged(nameof(StudentNo)); } }
public string Name { get => _name; set { _name = value; OnPropertyChanged(nameof(Name)); } }
public string ReceivedDate { get; set; } = "";

public string Status { get => _status; set { _status = value; OnPropertyChanged(nameof(Status)); OnPropertyChanged(nameof(StatusColor)); OnPropertyChanged(nameof(StatusForeColor)); } }
public string UkGrade { get => _ukGrade; set { _ukGrade = value; OnPropertyChanged(nameof(UkGrade)); OnPropertyChanged(nameof(GradeColor)); OnPropertyChanged(nameof(GradeForeColor)); } }

public IBrush StatusColor
{
    get
    {
        if (Status.Contains("✓")) return SolidColorBrush.Parse("#DCFCE7");
        if (Status.Contains("⚠️")) return SolidColorBrush.Parse("#FEF3C7");
        if (Status.Contains("⚙")) return SolidColorBrush.Parse("#DBEAFE");
        return SolidColorBrush.Parse("#F1F5F9");
    }
}

public IBrush StatusForeColor
{
    get
    {
        if (Status.Contains("✓")) return SolidColorBrush.Parse("#166534");
        if (Status.Contains("⚠️")) return SolidColorBrush.Parse("#92400E");
        if (Status.Contains("⚙")) return SolidColorBrush.Parse("#1E40AF");
        return SolidColorBrush.Parse("#475569");
    }
}

public IBrush GradeColor
{
    get
    {
        if (string.IsNullOrEmpty(UkGrade) || UkGrade == "??") return Brushes.Transparent;
        if (UkGrade == "1.0") return SolidColorBrush.Parse("#F0FDF4");
        if (UkGrade == "2.1") return SolidColorBrush.Parse("#EFF6FF");
        if (UkGrade == "2.2") return SolidColorBrush.Parse("#FFFBEB");
        return SolidColorBrush.Parse("#FEF2F2");
    }
}

public IBrush GradeForeColor
{
    get
    {
        if (string.IsNullOrEmpty(UkGrade) || UkGrade == "??") return SolidColorBrush.Parse("#94A3B8");
        if (UkGrade == "1.0") return SolidColorBrush.Parse("#15803D");
        if (UkGrade == "2.1") return SolidColorBrush.Parse("#1D4ED8");
        if (UkGrade == "2.2") return SolidColorBrush.Parse("#B45309");
        return SolidColorBrush.Parse("#B91C1C");
    }
}

public FontWeight GradeWeight => (string.IsNullOrEmpty(UkGrade) || UkGrade == "??") ? FontWeight.Normal : FontWeight.Bold;

public event PropertyChangedEventHandler? PropertyChanged;
protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

public partial class MainWindow : Window { private readonly ICsvService _csvService; private readonly IEquivalencyService _equivalencyService; private readonly IInstitutionMatchingService _matchingService; private readonly IRankingService _rankingService; private readonly IGradeClassificationService _gradeService;

private string _inTrayFilePath = string.Empty;
private string _appReportsFilePath = string.Empty;
private string _outputFolderPath = string.Empty;

private CancellationTokenSource? _cancellationTokenSource;
private bool _isProcessing = false;
private bool _rankingsLoaded = false;
public ObservableCollection<ProcessingItem> ProcessingItems { get; set; } = new ObservableCollection<ProcessingItem>();

[DllImport("winmm.dll")]
private static extern long mciSendString(string strCommand, StringBuilder? strReturn, int iReturnLength, IntPtr hwndCallback);

public MainWindow()
{
    InitializeComponent();
    ProcessingList.ItemsSource = ProcessingItems;
    _csvService = new CsvService();
    _equivalencyService = new EquivalencyService();
    _matchingService = new InstitutionMatchingService();
    _rankingService = new RankingService(_matchingService);
    _gradeService = new GradeClassificationService(_equivalencyService);
    
    SetVersion();
    InitializeDataAsync();
    
    BrowseInTrayButton.Click += BrowseInTrayButton_Click;
    BrowseAppReportsButton.Click += BrowseAppReportsButton_Click;
    BrowseOutputButton.Click += BrowseOutputButton_Click;
    ProcessButton.Click += ProcessButton_Click;
    ClearLogButton.Click += ClearLogButton_Click;
    ResetButton.Click += ResetButton_Click;
    ExitButton.Click += ExitButton_Click;
}

private async void InitializeDataAsync()
{
    StatusLabel.Text = "Loading ranking data in background...";
    try
    {
        _equivalencyService.LoadEquivalencies();
        await _rankingService.LoadRankingsAsync();
        _rankingsLoaded = true;
        Dispatcher.UIThread.Post(() => {
            StatusLabel.Text = "Ready";
            CheckReadyToProcess();
        });
    }
    catch (Exception ex) { LogStatus($"Load error: {ex.Message}"); }
}

private void CheckReadyToProcess()
{
    bool filesSelected = !string.IsNullOrEmpty(_inTrayFilePath) && 
                         !string.IsNullOrEmpty(_appReportsFilePath) && 
                         !string.IsNullOrEmpty(_outputFolderPath);
    ProcessButton.IsEnabled = filesSelected && _rankingsLoaded;
    if (!_rankingsLoaded) StatusLabel.Text = "Waiting for rankings...";
    else StatusLabel.Text = filesSelected ? "Ready to start" : "Waiting for files...";
}

private async Task ShowMessageBoxAsync(string title, string message)
{
    var okButton = new Button { 
        Content = "OK", 
        HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right, 
        Background = SolidColorBrush.Parse("#2563EB"), 
        Foreground = Brushes.White,
        Padding = new Thickness(20, 10)
    };

    var win = new Window { 
        Title = title, Width = 500, Height = 350, 
        WindowStartupLocation = WindowStartupLocation.CenterOwner, 
        SystemDecorations = SystemDecorations.BorderOnly, 
        ExtendClientAreaToDecorationsHint = true, 
        Content = new Border { 
            BorderBrush = Brushes.Gray, BorderThickness = new Thickness(1), Padding = new Thickness(30),
            Child = new StackPanel { 
                Spacing = 20, 
                Children = { 
                    new TextBlock { Text = title, FontWeight = FontWeight.Bold, FontSize = 20 }, 
                    new TextBlock { Text = message, TextWrapping = TextWrapping.Wrap, FontSize = 14 }, 
                    okButton 
                } 
            } 
        } 
    };

    okButton.Click += (s, e) => {
        win.Close();
        if (title == "Success" && !string.IsNullOrEmpty(_outputFolderPath) && Directory.Exists(_outputFolderPath)) {
            try {
                Process.Start(new ProcessStartInfo { FileName = _outputFolderPath, UseShellExecute = true });
            } catch { }
        }
    };
    await win.ShowDialog(this);
}

private async void ProcessButton_Click(object? sender, RoutedEventArgs e)
{
    if (_isProcessing)
    {
        _cancellationTokenSource?.Cancel();
        return;
    }

    _isProcessing = true;
    _cancellationTokenSource = new CancellationTokenSource();
    var token = _cancellationTokenSource.Token;
    ProcessButton.Content = "🛑 STOP";
    ProcessButton.Background = SolidColorBrush.Parse("#DC2626");

    try
    {
        await Task.Run(async () => {
            var inTrayRecords = _csvService.LoadInTrayRecords(_inTrayFilePath);
            var appRecords = _csvService.LoadApplicationRecords(_appReportsFilePath);
            
            await Dispatcher.UIThread.InvokeAsync(() => {
                ProcessingItems.Clear();
                foreach(var r in inTrayRecords) 
                    ProcessingItems.Add(new ProcessingItem { StudentNo = r.StudentNo, Name = r.Name, Status = "⏳ Pending" });
                FooterStatus.Text = $"0/{inTrayRecords.Count} Processed";
            });

            var outputRecords = new List<OutputRecord>();
            int current = 0;

            foreach (var inTray in inTrayRecords)
            {
                if (token.IsCancellationRequested) return;
                current++;

                var app = appRecords.FirstOrDefault(a => a.ApplicantID?.Trim() == inTray.StudentNo?.Trim());
                string classification = "??";

                if (app != null)
                {
                    classification = _gradeService.DetermineUKClassification(app.OverallGradeGPA ?? "", app.EquivalencyNote ?? "", app.CountryOfStudy ?? "", app.QualificationName ?? "");
                    
                    outputRecords.Add(new OutputRecord {
                        ReceivedDate = DateFormatter.FormatDate(inTray.ReceivedOn ?? ""),
                        DueDate = DateFormatter.CalculateDueDate(inTray.ReceivedOn ?? ""),
                        StudentNo = inTray.StudentNo,
                        Programme = ProgrammeMapping.GetCode(app.Programme ?? ""),
                        Forename = app.Forename,
                        Surname = app.Surname,
                        FeeStatus = app.FeeStatus,
                        QualificationName = app.QualificationName,
                        DegreeSubject = app.DegreeSubject,
                        InstitutionName = app.InstitutionName,
                        THERanking = _rankingService.GetRanking(app.InstitutionName ?? ""),
                        CountryOfStudy = app.CountryOfStudy,
                        EquivalencyNote = app.EquivalencyNote,
                        OverallGradeGPA = app.OverallGradeGPA,
                        DegreeStatus = app.GradeAchievedPending,
                        UKGrade = classification
                    });
                }

                await Dispatcher.UIThread.InvokeAsync(() => {
                    var item = ProcessingItems.FirstOrDefault(p => p.StudentNo == inTray.StudentNo);
                    if (item != null) { item.Status = app != null ? "✓ Done" : "⚠️ Missing"; item.UkGrade = classification; }
                    MainProgressBar.Value = (double)current / inTrayRecords.Count * 100;
                    FooterStatus.Text = $"{current}/{inTrayRecords.Count} Processed";
                });
            }

            _csvService.GenerateOutputFiles(outputRecords, _outputFolderPath);

            if (!token.IsCancellationRequested)
            {
                await Dispatcher.UIThread.InvokeAsync(async () => {
                    PlayConfirmationSound();
                    
                    int countFirst = outputRecords.Count(r => r.UKGrade == "1.0");
                    int countUpper = outputRecords.Count(r => r.UKGrade == "2.1");
                    int countLower = outputRecords.Count(r => r.UKGrade == "2.2");
                    int countThird = outputRecords.Count(r => r.UKGrade == "3.0");
                    int countOther = outputRecords.Count - (countFirst + countUpper + countLower + countThird);

                    string summaryList = $"{countFirst}\t(First Class)\n" +
                                         $"{countUpper}\t(Upper Second)\n" +
                                         $"{countLower}\t(Lower Second)";

                    if (countThird > 0) summaryList += $"\n{countThird}\t(Third Class)";
                    if (countOther > 0) summaryList += $"\n{countOther}\t(Other / Ungraded)";

                    string summary = $"Processed {outputRecords.Count} records:\n\n{summaryList}";

                    await ShowMessageBoxAsync("Success", 
                        $"Processing complete!\n\n{summary}\n\nExcel file(s) saved at:\n{_outputFolderPath}");
                });
            }
        }, token);
    }
    catch (Exception ex) { LogStatus($"Error: {ex.Message}"); }
    finally
    {
        _isProcessing = false;
        ProcessButton.Content = "Process Files";
        ProcessButton.Background = SolidColorBrush.Parse("#2563EB");
        CheckReadyToProcess();
    }
}

private void SetVersion() => VersionLabel.Text = "v1.0.0";

private async void BrowseInTrayButton_Click(object? sender, RoutedEventArgs e)
{
    var files = await GetTopLevel(this)!.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions { Title = "Select InTray" });
    if (files.Any()) {
        _inTrayFilePath = files[0].Path.LocalPath;
        InTrayFileLabel.Text = Path.GetFileName(_inTrayFilePath);
        CheckReadyToProcess();
    }
}

private async void BrowseAppReportsButton_Click(object? sender, RoutedEventArgs e)
{
    var files = await GetTopLevel(this)!.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions { Title = "Select App Reports" });
    if (files.Any()) {
        _appReportsFilePath = files[0].Path.LocalPath;
        AppReportsFileLabel.Text = Path.GetFileName(_appReportsFilePath);
        CheckReadyToProcess();
    }
}

private async void BrowseOutputButton_Click(object? sender, RoutedEventArgs e)
{
    var folders = await GetTopLevel(this)!.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions { Title = "Select Output" });
    if (folders.Count > 0) {
        _outputFolderPath = folders[0].Path.LocalPath;
        OutputFolderLabel.Text = _outputFolderPath;
        CheckReadyToProcess();
    }
}

private void PlayConfirmationSound()
{
    try
    {
        string soundPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "audio", "confirmed.mp3");

        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            if (!File.Exists(soundPath))
                soundPath = "/Users/xeon2035/Documents/LOCALDEV/admerger/audio/confirmed.mp3";

            if (File.Exists(soundPath))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "afplay",
                    Arguments = $"\"{soundPath}\"",
                    CreateNoWindow = true,
                    UseShellExecute = false
                });
            }
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            if (File.Exists(soundPath))
            {
                mciSendString($"open \"{soundPath}\" type mpegvideo alias confirm", null, 0, IntPtr.Zero);
                mciSendString("play confirm from 0", null, 0, IntPtr.Zero);
            }
        }
    }
    catch { }
}

private void LogStatus(string message) => Dispatcher.UIThread.Post(() => StatusLog.Text += $"[{DateTime.Now:HH:mm:ss}] {message}\n");
private void ClearLogButton_Click(object? sender, RoutedEventArgs e) => StatusLog.Text = string.Empty;
private void ExitButton_Click(object? sender, RoutedEventArgs e) => Environment.Exit(0);
private void ResetButton_Click(object? sender, RoutedEventArgs e) => CheckReadyToProcess();
}
