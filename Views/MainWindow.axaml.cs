// Views/MainWindow.axaml.cs

using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using Avalonia.Media; 
using System.Collections.ObjectModel; 
using ADMerger.Services;
using ADMerger.Models;
using ADMerger.Configuration;  
using ADMerger.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace ADMerger.Views;

// UPDATED ProcessingItem to support separate Status vs UK Grade
public class ProcessingItem : System.ComponentModel.INotifyPropertyChanged
{
    public string? StudentNo { get; set; }
    public string? Name { get; set; }
    public string? ReceivedDate { get; set; }
    
    // === STATUS COLUMN (Progress) ===
    private string _status = "Pending";
    public string Status 
    { 
        get => _status;
        set { _status = value; PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(Status))); OnStatusColorChanged(); } 
    }

    public IBrush StatusColor { get; private set; } = Brushes.Transparent;
    public IBrush StatusForeColor { get; private set; } = Brushes.Gray;

    // === UK GRADE COLUMN (Results) ===
    private string _ukGrade = ""; // Empty by default
    public string UkGrade
    {
        get => _ukGrade;
        set { _ukGrade = value; PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(UkGrade))); OnGradeColorChanged(); }
    }

    public IBrush GradeColor { get; private set; } = Brushes.Transparent;
    public IBrush GradeForeColor { get; private set; } = Brushes.Black;
    public FontWeight GradeWeight { get; private set; } = FontWeight.Normal;

    public event System.ComponentModel.PropertyChangedEventHandler? PropertyChanged;
    
    private void OnStatusColorChanged()
    {
        string s = Status.ToLower();
        if (s.Contains("done") || s.Contains("success")) {
            StatusColor = SolidColorBrush.Parse("#D1FAE5"); // Green
            StatusForeColor = SolidColorBrush.Parse("#059669"); 
        }
        else if (s.Contains("processing")) {
             StatusColor = SolidColorBrush.Parse("#DBEAFE"); // Blue
             StatusForeColor = SolidColorBrush.Parse("#2563EB"); 
        }
        else if (s.Contains("error") || s.Contains("missing") || s.Contains("stopped")) {
             StatusColor = SolidColorBrush.Parse("#FEE2E2"); // Red
             StatusForeColor = SolidColorBrush.Parse("#DC2626"); 
        }
        else {
             StatusColor = SolidColorBrush.Parse("#F1F5F9"); // Gray
             StatusForeColor = SolidColorBrush.Parse("#64748B"); 
        }
        PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(StatusColor)));
        PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(StatusForeColor)));
    }

    private void OnGradeColorChanged()
    {
        string g = UkGrade.ToLower();
        GradeWeight = FontWeight.Bold;
        
        // 1. High Achievement (1st, Distinction) -> Emerald Green
        if (g.Contains("1.0") || g.Contains("1st") || g.Contains("distinction")) 
        {
            GradeColor = SolidColorBrush.Parse("#D1FAE5"); 
            GradeForeColor = SolidColorBrush.Parse("#059669"); 
        }
        // 2. Good (2.1, Merit) -> Blue
        else if (g.Contains("2.1") || g.Contains("2:1") || g.Contains("merit")) 
        {
            GradeColor = SolidColorBrush.Parse("#DBEAFE"); 
            GradeForeColor = SolidColorBrush.Parse("#2563EB"); 
        }
        // 3. Okay (2.2, Pass) -> Orange/Amber
        else if (g.Contains("2.2") || g.Contains("2:2") || g.Contains("pass")) 
        {
            GradeColor = SolidColorBrush.Parse("#FEF3C7"); 
            GradeForeColor = SolidColorBrush.Parse("#D97706"); 
        }
        // 4. Low (3.0, Third) -> Rose/Red
        else if (g.Contains("3.0") || g.Contains("3rd")) 
        {
            GradeColor = SolidColorBrush.Parse("#FFE4E6"); 
            GradeForeColor = SolidColorBrush.Parse("#E11D48"); 
        }
        // 5. Masters (Generic) -> Violet/Purple
        else if (g.Contains("masters")) 
        {
            GradeColor = SolidColorBrush.Parse("#EDE9FE"); 
            GradeForeColor = SolidColorBrush.Parse("#7C3AED"); 
        }
        // Default -> Transparent/Black
        else 
        {
            GradeColor = Brushes.Transparent;
            GradeForeColor = Brushes.Black;
            GradeWeight = FontWeight.Normal;
        }

        PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(GradeColor)));
        PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(GradeForeColor)));
        PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(GradeWeight)));
    }
}

public partial class MainWindow : Window
{
    private readonly ICsvService _csvService;
    private readonly IEquivalencyService _equivalencyService;
    private readonly IInstitutionMatchingService _matchingService;
    private readonly IRankingService _rankingService;
    private readonly IGradeClassificationService _gradeService;
    
    private string _inTrayFilePath = string.Empty;
    private string _appReportsFilePath = string.Empty;
    private string _outputFolderPath = string.Empty;

    private CancellationTokenSource? _cancellationTokenSource;
    private bool _isProcessing = false;

    public ObservableCollection<ProcessingItem> ProcessingItems { get; set; } = new ObservableCollection<ProcessingItem>();

    [DllImport("winmm.dll")]
    private static extern long mciSendString(string strCommand, System.Text.StringBuilder? strReturn, int iReturnLength, IntPtr hwndCallback);

    public MainWindow()
    {
        InitializeComponent();
        ProcessingList.ItemsSource = ProcessingItems;

        _csvService = new CsvService();
        _equivalencyService = new EquivalencyService();
        _matchingService = new InstitutionMatchingService();
        _rankingService = new RankingService(_matchingService);
        _gradeService = new GradeClassificationService(_equivalencyService);
        
        LoadRankingsAndEquivalencies();
        SetVersion();
        
        BrowseInTrayButton.Click += BrowseInTrayButton_Click;
        BrowseAppReportsButton.Click += BrowseAppReportsButton_Click;
        BrowseOutputButton.Click += BrowseOutputButton_Click;
        ProcessButton.Click += ProcessButton_Click;
        ClearLogButton.Click += ClearLogButton_Click;
        ResetButton.Click += ResetButton_Click;
        ExitButton.Click += ExitButton_Click;
    }
    
    private void SetVersion()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var infoVersionAttr = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>();
        string version = infoVersionAttr?.InformationalVersion ?? "1.0.0";
        if (string.IsNullOrEmpty(version)) version = assembly.GetName().Version?.ToString() ?? "1.0.0";
        if (version.Contains('+')) version = version.Split('+')[0];
        var parts = version.Split('.');
        if (parts.Length >= 3 && parts[2].Length >= 2) version = $"{parts[0]}.{parts[1]}.{parts[2].Substring(0, 2)}";
        VersionLabel.Text = $"v{version}";
    }
    
    private void LoadRankingsAndEquivalencies()
    {
        try
        {
            _rankingService.LoadRankings();
            LogStatus($"Loaded {_rankingService.Count} THE World University Rankings");
            _equivalencyService.LoadEquivalencies();
            LogStatus($"Loaded {_equivalencyService.Count} degree equivalencies");
        }
        catch (Exception ex)
        {
            LogStatus($"Warning: Could not load reference data: {ex.Message}");
        }
    }
    
    private bool ValidateFileSignature(string path)
    {
        try
        {
            if (!File.Exists(path)) return false;
            var extension = Path.GetExtension(path).ToLower();
            var buffer = new byte[4];
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                if (fs.Length == 0) return false;
                fs.Read(buffer, 0, 4);
            }

            bool isZip = (buffer[0] == 0x50 && buffer[1] == 0x4B);

            if (extension == ".xlsx" || extension == ".xls")
            {
                if (!isZip && extension == ".xlsx") 
                {
                    LogStatus($"ERROR: {Path.GetFileName(path)} is not a valid Excel file.");
                    return false;
                }
            }
            else if (extension == ".csv" || extension == ".txt")
            {
                if (isZip)
                {
                    LogStatus($"ERROR: {Path.GetFileName(path)} looks like an Excel file but is named .csv.");
                    return false;
                }
                const int checkSize = 1024;
                var checkBuffer = new byte[checkSize];
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    int read = fs.Read(checkBuffer, 0, checkSize);
                    for (int i = 0; i < read; i++)
                    {
                        if (checkBuffer[i] == 0) 
                        {
                            LogStatus($"ERROR: {Path.GetFileName(path)} appears to be a binary file.");
                            return false;
                        }
                    }
                }
            }
            return true;
        }
        catch { return false; }
    }

    private async void BrowseInTrayButton_Click(object? sender, RoutedEventArgs e)
    {
        var topLevel = GetTopLevel(this);
        if (topLevel == null) return;
        
        var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Select InTray File",
            AllowMultiple = false,
            FileTypeFilter = new[] { new FilePickerFileType("Excel or CSV") { Patterns = new[] { "*.xlsx", "*.csv" } } }
        });

        if (files.Count > 0)
        {
            var path = files[0].Path.LocalPath;
            if (!ValidateFileSignature(path)) {
                LogStatus("Error: Invalid InTray file format.");
                StatusLabel.Text = "Invalid File Format";
                StatusLabel.Foreground = Brushes.Red;
                return;
            }

            _inTrayFilePath = path;
            InTrayFileLabel.Text = Path.GetFileName(_inTrayFilePath);
            LogStatus($"Selected InTray file: {Path.GetFileName(_inTrayFilePath)}");
            CheckReadyToProcess();
            
            try {
                var records = _csvService.LoadInTrayRecords(_inTrayFilePath);
                ProcessingItems.Clear();
                foreach(var r in records.Take(50)) { 
                    ProcessingItems.Add(new ProcessingItem { StudentNo = r.StudentNo, Name = r.Name, ReceivedDate = r.ReceivedOn, Status = "⏳ Pending", UkGrade = "" });
                }
                if (records.Count > 50) LogStatus($"...and {records.Count - 50} more records.");
                EmptyStatePanel.IsVisible = false;
                FooterStatus.Text = $"{records.Count} Records waiting";
                StatusLabel.Foreground = Brushes.Black; 
            } catch { }
        }
    }
    
    private async void BrowseAppReportsButton_Click(object? sender, RoutedEventArgs e)
    {
        var topLevel = GetTopLevel(this);
        if (topLevel == null) return;
        
        var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Select Application Reports",
            AllowMultiple = false,
            FileTypeFilter = new[] { new FilePickerFileType("Excel or CSV") { Patterns = new[] { "*.xlsx", "*.csv" } } }
        });
        if (files.Count > 0)
        {
            var path = files[0].Path.LocalPath;
            if (!ValidateFileSignature(path)) {
                LogStatus("Error: Invalid App Reports file format.");
                StatusLabel.Text = "Invalid File Format";
                StatusLabel.Foreground = Brushes.Red;
                return;
            }

            _appReportsFilePath = path;
            AppReportsFileLabel.Text = Path.GetFileName(_appReportsFilePath);
            LogStatus($"Selected App Reports: {Path.GetFileName(_appReportsFilePath)}");
            CheckReadyToProcess();
            StatusLabel.Foreground = Brushes.Black; 
        }
    }
    
    private async void BrowseOutputButton_Click(object? sender, RoutedEventArgs e)
    {
        var topLevel = GetTopLevel(this);
        if (topLevel == null) return;
        var folders = await topLevel.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions { Title = "Select Output Folder" });
        if (folders.Count > 0)
        {
            _outputFolderPath = folders[0].Path.LocalPath;
            OutputFolderLabel.Text = _outputFolderPath;
            CheckReadyToProcess();
        }
    }
    
    private void CheckReadyToProcess()
    {
        bool ready = !string.IsNullOrEmpty(_inTrayFilePath) && !string.IsNullOrEmpty(_appReportsFilePath) && !string.IsNullOrEmpty(_outputFolderPath);
        ProcessButton.IsEnabled = ready;
        StatusLabel.Text = ready ? "Ready to start processing" : "Waiting for files...";
        if (ready) StatusLabel.Foreground = Brushes.Black;
    }
    
    private async void ProcessButton_Click(object? sender, RoutedEventArgs e)
    {
        if (_isProcessing)
        {
            _cancellationTokenSource?.Cancel();
            ProcessButton.Content = "Stopping...";
            ProcessButton.IsEnabled = false; 
            LogStatus("User requested stop.");
            return;
        }

        if (!ValidateFileSignature(_inTrayFilePath) || !ValidateFileSignature(_appReportsFilePath))
        {
             StatusLabel.Text = "ERROR: Make sure you're importing the right spreadsheets.";
             StatusLabel.Foreground = Brushes.Red;
             LogStatus("Process aborted: Invalid file signatures.");
             return; 
        }

        _isProcessing = true;
        _cancellationTokenSource = new CancellationTokenSource();
        var token = _cancellationTokenSource.Token;

        ProcessButton.Content = "🛑 STOP"; 
        ProcessButton.Background = SolidColorBrush.Parse("#DC2626"); 
        BrowseInTrayButton.IsEnabled = false;
        BrowseAppReportsButton.IsEnabled = false;
        BrowseOutputButton.IsEnabled = false;
        MainProgressBar.Value = 0;
        StatusLabel.Foreground = Brushes.Black;

        try
        {
            LogStatus("=== Starting Processing ===");
            StatusLabel.Text = "Reading files...";
            
            await Task.Run(async () => {
                if (token.IsCancellationRequested) return;

                var inTrayRecords = _csvService.LoadInTrayRecords(_inTrayFilePath);
                if (token.IsCancellationRequested) return;
                var appRecords = _csvService.LoadApplicationRecords(_appReportsFilePath);
                
                await Dispatcher.UIThread.InvokeAsync(() => {
                    ProcessingItems.Clear();
                    EmptyStatePanel.IsVisible = false;
                    foreach(var r in inTrayRecords) 
                        ProcessingItems.Add(new ProcessingItem { StudentNo = r.StudentNo, Name = r.Name, ReceivedDate = r.ReceivedOn, Status = "⏳ Pending", UkGrade = "" });
                });

                var outputRecords = new List<OutputRecord>();
                int total = inTrayRecords.Count;
                int current = 0;

                foreach (var inTray in inTrayRecords)
                {
                    if (token.IsCancellationRequested) {
                        await Dispatcher.UIThread.InvokeAsync(() => StatusLabel.Text = "Stopped");
                        return;
                    }

                    current++;
                    
                    var studentNo = inTray.StudentNo?.Trim();
                    var app = appRecords.FirstOrDefault(a => a.ApplicantID?.Trim() == studentNo);

                    string newStatus = "⚙ Processing...";
                    string newGrade = "";

                    if (app == null)
                    {
                        newStatus = "⚠️ Missing";
                        newGrade = "N/A";
                    }
                    else
                    {
                        var programmeCode = ProgrammeMapping.GetCode(app.Programme ?? "");
                        
                        var ukGrade = _gradeService.DetermineUKClassification(
                            app.OverallGradeGPA ?? "", 
                            app.EquivalencyNote ?? "", 
                            app.CountryOfStudy ?? "",
                            app.QualificationName ?? "");

                        var theRanking = _rankingService.GetRanking(app.InstitutionName ?? "");
                        
                        outputRecords.Add(new OutputRecord
                        {
                            ReceivedDate = DateFormatter.FormatDate(inTray.ReceivedOn ?? ""),
                            DueDate = DateFormatter.CalculateDueDate(inTray.ReceivedOn ?? ""),
                            StudentNo = inTray.StudentNo,
                            Programme = programmeCode,
                            Forename = app.Forename,
                            Surname = app.Surname,
                            FeeStatus = app.FeeStatus,
                            QualificationName = app.QualificationName,
                            DegreeSubject = app.DegreeSubject,
                            InstitutionName = app.InstitutionName,
                            THERanking = theRanking,
                            CountryOfStudy = app.CountryOfStudy,
                            EquivalencyNote = app.EquivalencyNote,
                            OverallGradeGPA = app.OverallGradeGPA,
                            DegreeStatus = app.GradeAchievedPending,
                            UKGrade = ukGrade
                        });

                        newStatus = "✓ Done";
                        newGrade = ukGrade;
                    }

                    await Dispatcher.UIThread.InvokeAsync(() => {
                        var uiItem = ProcessingItems.FirstOrDefault(p => p.StudentNo == inTray.StudentNo);
                        if (uiItem != null) 
                        {
                            uiItem.Status = newStatus;   
                            uiItem.UkGrade = newGrade;   
                        }
                        MainProgressBar.Value = (double)current / total * 100;
                        FooterStatus.Text = $"{current}/{total} Processed";
                    });

                    // 50ms delay for visual effect
                    await Task.Delay(50); 
                }

                if (outputRecords.Count == 0)
                {
                    await Dispatcher.UIThread.InvokeAsync(async () => {
                         StatusLabel.Text = "No matches found";
                         await ShowMessageBoxAsync("No Matches", "No matches found between files.");
                    });
                    return;
                }
                
                if (token.IsCancellationRequested) return;

                await Dispatcher.UIThread.InvokeAsync(() => StatusLabel.Text = "Generating Excel files...");
                var outputPaths = _csvService.GenerateOutputFiles(outputRecords, _outputFolderPath);
                
                PlaySuccessSound();

                await Dispatcher.UIThread.InvokeAsync(async () => {
                    StatusLabel.Text = "Complete!";
                    MainProgressBar.Value = 100;
                    FooterStatus.Text = "Finished";
                    foreach (var path in outputPaths) LogStatus($"Generated: {Path.GetFileName(path)}");
                    
                    await ShowMessageBoxAsync("Success", 
                        $"Processing complete!\n\nExcel file(s) saved at:\n{_outputFolderPath}");
                });
            }, token); 
        }
        catch (Exception ex)
        {
            LogStatus($"ERROR: {ex.Message}");
            StatusLabel.Text = "ERROR: Make sure you're importing the right spreadsheets.";
            StatusLabel.Foreground = Brushes.Red;
        }
        finally
        {
            _isProcessing = false;
            _cancellationTokenSource?.Dispose();
            _cancellationTokenSource = null;
            ProcessButton.Content = "Process Files";
            ProcessButton.Background = SolidColorBrush.Parse("#2563EB");
            ProcessButton.IsEnabled = true;
            BrowseInTrayButton.IsEnabled = true;
            BrowseAppReportsButton.IsEnabled = true;
            BrowseOutputButton.IsEnabled = true;
        }
    }

    private void PlaySuccessSound()
    {
        try {
            string tempPath = Path.Combine(Path.GetTempPath(), "admerger_confirmed.mp3");
            if (!File.Exists(tempPath)) {
                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ADMerger.audio.confirmed.mp3"))
                using (var fileStream = File.Create(tempPath)) stream?.CopyTo(fileStream);
            }
            if (File.Exists(tempPath)) {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) System.Diagnostics.Process.Start("afplay", $"\"{tempPath}\"");
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
                    mciSendString($"open \"{tempPath}\" type mpegvideo alias MyMp3", null, 0, IntPtr.Zero);
                    mciSendString("play MyMp3", null, 0, IntPtr.Zero);
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) System.Diagnostics.Process.Start("paplay", $"\"{tempPath}\"");
            }
        } catch { }
    }
    
    private void ClearLogButton_Click(object? sender, RoutedEventArgs e)
    {
        StatusLog.Text = string.Empty;
        StatusLabel.Foreground = Brushes.Black;
    }
    
    private void ResetButton_Click(object? sender, RoutedEventArgs e)
    {
        if (_isProcessing) _cancellationTokenSource?.Cancel();
        _inTrayFilePath = string.Empty;
        _appReportsFilePath = string.Empty;
        _outputFolderPath = string.Empty;
        InTrayFileLabel.Text = "No file selected";
        AppReportsFileLabel.Text = "No file selected";
        OutputFolderLabel.Text = "No folder selected";
        StatusLog.Text = string.Empty;
        ProcessingItems.Clear();
        EmptyStatePanel.IsVisible = true;
        MainProgressBar.Value = 0;
        FooterStatus.Text = "0/0 Records";
        StatusLabel.Text = "Waiting for files...";
        StatusLabel.Foreground = Brushes.Black;
        CheckReadyToProcess();
    }
    
    private void ExitButton_Click(object? sender, RoutedEventArgs e)
    {
        Environment.Exit(0);
    }
    
    private void LogStatus(string message)
    {
        Dispatcher.UIThread.Post(() => {
            StatusLog.Text += $"[{DateTime.Now:HH:mm:ss}] {message}\n";
            StatusLog.CaretIndex = StatusLog.Text?.Length ?? 0;
        });
    }
    
    private async System.Threading.Tasks.Task ShowMessageBoxAsync(string title, string message)
    {
        var win = new Window {
            Title = title, Width = 400, Height = 200, WindowStartupLocation = WindowStartupLocation.CenterOwner,
            SystemDecorations = SystemDecorations.BorderOnly, ExtendClientAreaToDecorationsHint = true,
            Content = new Border {
                BorderBrush = Brushes.Gray, BorderThickness = new Thickness(1), Padding = new Thickness(20),
                Child = new StackPanel {
                    Spacing = 20,
                    Children = {
                        new TextBlock { Text = title, FontWeight = FontWeight.Bold, FontSize = 18 },
                        new TextBlock { Text = message, TextWrapping = TextWrapping.Wrap },
                        new Button { Content = "OK", HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right, Background = SolidColorBrush.Parse("#2563EB"), Foreground = Brushes.White }
                    }
                }
            }
        };
        ((Button)((StackPanel)((Border)win.Content).Child).Children[2]).Click += (s, e) => win.Close();
        await win.ShowDialog(this);
    }
}
