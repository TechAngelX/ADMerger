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

// Simple UI Model for the list
public class ProcessingItem : System.ComponentModel.INotifyPropertyChanged
{
    public string? StudentNo { get; set; }
    public string? Name { get; set; }
    public string? ReceivedDate { get; set; }
    
    private string _status = "Pending";
    public string Status 
    { 
        get => _status;
        set { _status = value; PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(Status))); OnColorsChanged(); } 
    }

    // Dynamic Colors based on Status
    public IBrush StatusColor { get; private set; } = Brushes.LightGray;
    public IBrush StatusForeColor { get; private set; } = Brushes.Gray;

    public event System.ComponentModel.PropertyChangedEventHandler? PropertyChanged;
    private void OnColorsChanged()
    {
        if (Status.Contains("Done") || Status.Contains("Success")) {
            StatusColor = SolidColorBrush.Parse("#D1FAE5"); // Light Green
            StatusForeColor = SolidColorBrush.Parse("#059669"); // Dark Green
        }
        else if (Status.Contains("Processing")) {
             StatusColor = SolidColorBrush.Parse("#DBEAFE"); // Light Blue
             StatusForeColor = SolidColorBrush.Parse("#2563EB"); // Blue
        }
        else if (Status.Contains("Error") || Status.Contains("Missing") || Status.Contains("Stopped")) {
             StatusColor = SolidColorBrush.Parse("#FEE2E2"); // Light Red
             StatusForeColor = SolidColorBrush.Parse("#DC2626"); // Red
        }
        else {
             StatusColor = SolidColorBrush.Parse("#F1F5F9"); // Gray
             StatusForeColor = SolidColorBrush.Parse("#64748B"); // Dark Gray
        }
        PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(StatusColor)));
        PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(StatusForeColor)));
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

    // Cancellation support
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
        if (string.IsNullOrEmpty(version))
        {
            version = assembly.GetName().Version?.ToString() ?? "1.0.0";
        }

        if (version.Contains('+'))
        {
            version = version.Split('+')[0];
        }

        var parts = version.Split('.');
        if (parts.Length >= 3 && parts[2].Length >= 2)
        {
            version = $"{parts[0]}.{parts[1]}.{parts[2].Substring(0, 2)}";
        }
        
        VersionLabel.Text = $"v{version}";
    }
    
    private void LoadRankingsAndEquivalencies()
    {
        try
        {
            _rankingService.LoadRankings();
            var rankingCount = _rankingService.Count;
            LogStatus($"Loaded {rankingCount} THE World University Rankings");
            
            _equivalencyService.LoadEquivalencies();
            var equivCount = _equivalencyService.Count;
            LogStatus($"Loaded {equivCount} degree equivalencies");
        }
        catch (Exception ex)
        {
            LogStatus($"Warning: Could not load reference data: {ex.Message}");
        }
    }
    
    private async void BrowseInTrayButton_Click(object? sender, RoutedEventArgs e)
    {
        var topLevel = GetTopLevel(this);
        if (topLevel == null) return;
        
        var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Select InTray File",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("Excel or CSV Files")
                {
                    Patterns = new[] { "*.xlsx", "*.csv" },
                    AppleUniformTypeIdentifiers = new[] { "public.comma-separated-values-text", "org.openxmlformats.spreadsheetml.sheet" },
                    MimeTypes = new[] { "text/csv", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
                }
            }
        });

        if (files.Count > 0)
        {
            _inTrayFilePath = files[0].Path.LocalPath;
            InTrayFileLabel.Text = Path.GetFileName(_inTrayFilePath);
            LogStatus($"Selected InTray file: {Path.GetFileName(_inTrayFilePath)}");
            CheckReadyToProcess();
            
            // Try-catch block to prevent crash on bad file selection immediately
            try {
                var records = _csvService.LoadInTrayRecords(_inTrayFilePath);
                ProcessingItems.Clear();
                foreach(var r in records) {
                    ProcessingItems.Add(new ProcessingItem { 
                        StudentNo = r.StudentNo, 
                        Name = r.Name, 
                        ReceivedDate = r.ReceivedOn,
                        Status = "⏳ Pending"
                    });
                }
                EmptyStatePanel.IsVisible = false;
                FooterStatus.Text = $"{ProcessingItems.Count} Records waiting";
            } 
            catch (Exception ex) {
                LogStatus($"Error previewing file: {ex.Message}");
                StatusLabel.Text = "File Error (Check Log)";
            }
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
            FileTypeFilter = new[]
            {
                new FilePickerFileType("Excel or CSV Files")
                {
                    Patterns = new[] { "*.xlsx", "*.csv" },
                    AppleUniformTypeIdentifiers = new[] { "public.comma-separated-values-text", "org.openxmlformats.spreadsheetml.sheet" },
                    MimeTypes = new[] { "text/csv", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
                }
            }
        });
        if (files.Count > 0)
        {
            _appReportsFilePath = files[0].Path.LocalPath;
            AppReportsFileLabel.Text = Path.GetFileName(_appReportsFilePath);
            LogStatus($"Selected Application Reports file: {Path.GetFileName(_appReportsFilePath)}");
            CheckReadyToProcess();
        }
    }
    
    private async void BrowseOutputButton_Click(object? sender, RoutedEventArgs e)
    {
        var topLevel = GetTopLevel(this);
        if (topLevel == null) return;
        
        var folders = await topLevel.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
        {
            Title = "Select Output Folder",
            AllowMultiple = false
        });
        if (folders.Count > 0)
        {
            _outputFolderPath = folders[0].Path.LocalPath;
            OutputFolderLabel.Text = _outputFolderPath;
            LogStatus($"Selected output folder: {_outputFolderPath}");
            CheckReadyToProcess();
        }
    }
    
    private void CheckReadyToProcess()
    {
        bool ready = !string.IsNullOrEmpty(_inTrayFilePath) &&
                     !string.IsNullOrEmpty(_appReportsFilePath) &&
                     !string.IsNullOrEmpty(_outputFolderPath);
        ProcessButton.IsEnabled = ready;
        
        if (ready)
        {
            StatusLabel.Text = "Ready to start processing";
        }
        else
        {
            StatusLabel.Text = "Waiting for files...";
        }
    }
    
    private async void ProcessButton_Click(object? sender, RoutedEventArgs e)
    {
        // === CANCELLATION LOGIC ===
        if (_isProcessing)
        {
            _cancellationTokenSource?.Cancel();
            ProcessButton.Content = "Stopping...";
            ProcessButton.IsEnabled = false; // Prevent double clicks
            StatusLabel.Text = "Stopping operation...";
            LogStatus("User requested cancellation.");
            return;
        }

        // === START PROCESSING ===
        _isProcessing = true;
        _cancellationTokenSource = new CancellationTokenSource();
        var token = _cancellationTokenSource.Token;

        ProcessButton.Content = "🛑 STOP"; // Change button to Stop
        ProcessButton.Background = SolidColorBrush.Parse("#DC2626"); // Red color
        
        BrowseInTrayButton.IsEnabled = false;
        BrowseAppReportsButton.IsEnabled = false;
        BrowseOutputButton.IsEnabled = false;
        MainProgressBar.Value = 0;

        try
        {
            LogStatus("=== Starting Processing ===");
            StatusLabel.Text = "Reading files (this may take a moment)...";
            
            await Task.Run(async () => {
                
                // Check cancellation before heavy lifting
                if (token.IsCancellationRequested) return;

                var inTrayRecords = _csvService.LoadInTrayRecords(_inTrayFilePath);
                
                if (token.IsCancellationRequested) return;
                var appRecords = _csvService.LoadApplicationRecords(_appReportsFilePath);
                
                await Dispatcher.UIThread.InvokeAsync(() => {
                    ProcessingItems.Clear();
                    EmptyStatePanel.IsVisible = false;
                    foreach(var r in inTrayRecords) {
                        ProcessingItems.Add(new ProcessingItem { 
                            StudentNo = r.StudentNo, 
                            Name = r.Name, 
                            ReceivedDate = r.ReceivedOn,
                            Status = "⏳ Pending"
                        });
                    }
                });

                var outputRecords = new List<OutputRecord>();
                int total = inTrayRecords.Count;
                int current = 0;

                foreach (var inTray in inTrayRecords)
                {
                    // === CHECK CANCELLATION INSIDE LOOP ===
                    if (token.IsCancellationRequested)
                    {
                        await Dispatcher.UIThread.InvokeAsync(() => {
                            StatusLabel.Text = "Operation Stopped";
                            LogStatus("Processing stopped by user.");
                        });
                        return; // Break out of the background task
                    }

                    current++;
                    await Dispatcher.UIThread.InvokeAsync(() => {
                        var uiItem = ProcessingItems.FirstOrDefault(p => p.StudentNo == inTray.StudentNo);
                        if (uiItem != null) uiItem.Status = "⚙ Processing...";
                        MainProgressBar.Value = (double)current / total * 100;
                        FooterStatus.Text = $"{current}/{total} Records";
                    });

                    var studentNo = inTray.StudentNo?.Trim();
                    var app = appRecords.FirstOrDefault(a => a.ApplicantID?.Trim() == studentNo);

                    if (app == null)
                    {
                        LogStatus($"Warning: No application record found for {inTray.StudentNo}");
                        await Dispatcher.UIThread.InvokeAsync(() => {
                             var uiItem = ProcessingItems.FirstOrDefault(p => p.StudentNo == inTray.StudentNo);
                             if (uiItem != null) uiItem.Status = "⚠️ Missing Data";
                        });
                        continue;
                    }
                    
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

                    await Dispatcher.UIThread.InvokeAsync(() => {
                         var uiItem = ProcessingItems.FirstOrDefault(p => p.StudentNo == inTray.StudentNo);
                         if (uiItem != null) uiItem.Status = "✓ Done";
                    });

                    // Small delay to keep UI responsive
                    await Task.Delay(5); 
                }

                if (outputRecords.Count == 0)
                {
                    await Dispatcher.UIThread.InvokeAsync(async () => {
                         StatusLabel.Text = "No matches found";
                         await ShowMessageBoxAsync("No Matches", "No matches found between files.");
                    });
                    return;
                }
                
                // Check cancellation before saving
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
        catch (OperationCanceledException)
        {
             LogStatus("Operation cancelled.");
             StatusLabel.Text = "Cancelled";
        }
        catch (Exception ex)
        {
            LogStatus($"ERROR: {ex.Message}");
            StatusLabel.Text = "Error occurred";
            await ShowMessageBoxAsync("Error", $"Processing failed:\n\n{ex.Message}");
        }
        finally
        {
            // === RESET UI STATE ===
            _isProcessing = false;
            if (_cancellationTokenSource != null)
            {
                _cancellationTokenSource.Dispose();
                _cancellationTokenSource = null;
            }

            ProcessButton.Content = "Process Files"; // Reset Text
            ProcessButton.Background = SolidColorBrush.Parse("#2563EB"); // Reset Color (Blue)
            ProcessButton.IsEnabled = true;
            BrowseInTrayButton.IsEnabled = true;
            BrowseAppReportsButton.IsEnabled = true;
            BrowseOutputButton.IsEnabled = true;
        }
    }

    private void PlaySuccessSound()
    {
        try
        {
            string tempPath = Path.Combine(Path.GetTempPath(), "admerger_confirmed.mp3");
            if (!File.Exists(tempPath))
            {
                var assembly = Assembly.GetExecutingAssembly();
                var resourceName = "ADMerger.audio.confirmed.mp3";
                using (var stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream != null)
                    {
                        using (var fileStream = File.Create(tempPath))
                        {
                            stream.CopyTo(fileStream);
                        }
                    }
                }
            }

            if (File.Exists(tempPath))
            {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                {
                    System.Diagnostics.Process.Start("afplay", $"\"{tempPath}\"");
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    mciSendString($"open \"{tempPath}\" type mpegvideo alias MyMp3", null, 0, IntPtr.Zero);
                    mciSendString("play MyMp3", null, 0, IntPtr.Zero);
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                {
                    System.Diagnostics.Process.Start("paplay", $"\"{tempPath}\"");
                }
            }
        }
        catch { }
    }
    
    private void ClearLogButton_Click(object? sender, RoutedEventArgs e)
    {
        StatusLog.Text = string.Empty;
        LogStatus("Log cleared");
    }
    
    private void ResetButton_Click(object? sender, RoutedEventArgs e)
    {
        // Cancel any running process first
        if (_isProcessing)
        {
            _cancellationTokenSource?.Cancel();
        }

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
        
        CheckReadyToProcess();
        LogStatus("Application reset");
    }
    
    private void ExitButton_Click(object? sender, RoutedEventArgs e)
    {
        // HARD EXIT: Kills the process immediately. 
        // This is useful if the app is stuck reading a bad file.
        Environment.Exit(0);
    }
    
    private void LogStatus(string message)
    {
        Dispatcher.UIThread.Post(() =>
        {
            var timestamp = DateTime.Now.ToString("HH:mm:ss");
            StatusLog.Text += $"[{timestamp}] {message}\n";
            StatusLog.CaretIndex = StatusLog.Text?.Length ?? 0;
        });
    }
    
    private async System.Threading.Tasks.Task ShowMessageBoxAsync(string title, string message)
    {
        var messageBox = new Window
        {
            Title = title,
            Width = 450,
            Height = 250,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            CanResize = false,
            SystemDecorations = SystemDecorations.BorderOnly,
            ExtendClientAreaToDecorationsHint = true,
            Background = SolidColorBrush.Parse("#FFFFFF"),
            Content = new Border 
            {
                BorderBrush = SolidColorBrush.Parse("#E2E8F0"),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(20),
                Child = new StackPanel
                {
                    Spacing = 20,
                    Children =
                    {
                        new TextBlock { Text = title, FontWeight = FontWeight.Bold, FontSize = 18, Foreground = SolidColorBrush.Parse("#1E293B") },
                        new TextBlock 
                        { 
                            Text = message, 
                            TextWrapping = TextWrapping.Wrap,
                            FontSize = 14,
                            Foreground = SolidColorBrush.Parse("#475569")
                        },
                        new Button
                        {
                            Content = "OK",
                            HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right,
                            Padding = new Thickness(30, 10),
                            Background = SolidColorBrush.Parse("#2563EB"),
                            Foreground = Brushes.White,
                            CornerRadius = new CornerRadius(6)
                        }
                    }
                }
            }
        };
        
        var button = ((StackPanel)((Border)messageBox.Content).Child).Children.OfType<Button>().First();
        button.Click += (s, e) => messageBox.Close();
        
        await messageBox.ShowDialog(this);
    }
}
