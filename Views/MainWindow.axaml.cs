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
using System.Diagnostics;
using System.ComponentModel;
using System.Text;
using ADMerger.data;

namespace ADMerger.Views;

public class ProcessingItem : INotifyPropertyChanged
{
    private string _status = "";
    private string _ukGrade = "";
    private string _studentNo = "";
    private string _name = "";

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

public class AdditionalField
{
    public string DisplayName { get; set; } = "";
    public string PropertyName { get; set; } = "";
    public string Category { get; set; } = "";
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
    private bool _rankingsLoaded = false;
    public ObservableCollection<ProcessingItem> ProcessingItems { get; set; } = new ObservableCollection<ProcessingItem>();

    private readonly List<AdditionalField> _additionalFields = new List<AdditionalField>();

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
        InitializeAdditionalFields();
        InitializeDataAsync();

        BrowseInTrayButton.Click += BrowseInTrayButton_Click;
        BrowseAppReportsButton.Click += BrowseAppReportsButton_Click;
        BrowseOutputButton.Click += BrowseOutputButton_Click;
        ProcessButton.Click += ProcessButton_Click;
        ClearLogButton.Click += ClearLogButton_Click;
        ResetButton.Click += ResetButton_Click;
        ExitButton.Click += ExitButton_Click;
    }

    private void InitializeAdditionalFields()
    {
        // Personal Information fields
        var personalFields = new List<AdditionalField>
        {
            new AdditionalField { DisplayName = "Known as", PropertyName = "KnownAs", Category = "Personal" },
            new AdditionalField { DisplayName = "Email address", PropertyName = "EmailAddress", Category = "Personal" },
            new AdditionalField { DisplayName = "Gender", PropertyName = "Gender", Category = "Personal" },
            new AdditionalField { DisplayName = "Date of Birth", PropertyName = "DateOfBirth", Category = "Personal" },
            new AdditionalField { DisplayName = "Country of Nationality", PropertyName = "CountryOfNationality", Category = "Personal" }
        };

        // Application Details fields
        var applicationFields = new List<AdditionalField>
        {
            new AdditionalField { DisplayName = "Mode of attendance", PropertyName = "ModeOfAttendance", Category = "Application" },
            new AdditionalField { DisplayName = "Location", PropertyName = "Location", Category = "Application" },
            new AdditionalField { DisplayName = "State", PropertyName = "State", Category = "Application" },
            new AdditionalField { DisplayName = "Academic year", PropertyName = "AcademicYear", Category = "Application" },
            new AdditionalField { DisplayName = "Tag", PropertyName = "Tag", Category = "Application" },
            new AdditionalField { DisplayName = "Qualification end date", PropertyName = "QualificationEndDate", Category = "Application" },
            new AdditionalField { DisplayName = "Total Mark equivalency", PropertyName = "TotalMarkEquivalency", Category = "Application" },
            new AdditionalField { DisplayName = "ELP type", PropertyName = "ELPType", Category = "Application" },
            new AdditionalField { DisplayName = "ELP verification status", PropertyName = "ELPVerificationStatus", Category = "Application" }
        };

        // Decision Tracking fields
        var decisionFields = new List<AdditionalField>
        {
            new AdditionalField { DisplayName = "Admissions referral note", PropertyName = "AdmissionsReferralNote", Category = "Decision" },
            new AdditionalField { DisplayName = "Admissions referred to dept date", PropertyName = "AdmissionsReferredToDepartmentDate", Category = "Decision" },
            new AdditionalField { DisplayName = "Department recommended decision", PropertyName = "DepartmentRecommendedDecision", Category = "Decision" },
            new AdditionalField { DisplayName = "Department recommended decision date", PropertyName = "DepartmentRecommendedDecisionDate", Category = "Decision" },
            new AdditionalField { DisplayName = "Deposit due date", PropertyName = "DepositDueDate", Category = "Decision" },
            new AdditionalField { DisplayName = "Deposit payment status", PropertyName = "DepositPaymentStatus", Category = "Decision" },
            new AdditionalField { DisplayName = "Initial decision", PropertyName = "InitialDecision", Category = "Decision" },
            new AdditionalField { DisplayName = "Initial Decision date", PropertyName = "InitialDecisionDate", Category = "Decision" },
            new AdditionalField { DisplayName = "Decision/Response", PropertyName = "DecisionResponse", Category = "Decision" },
            new AdditionalField { DisplayName = "Reply by date", PropertyName = "ReplyByDate", Category = "Decision" }
        };

        _additionalFields.AddRange(personalFields);
        _additionalFields.AddRange(applicationFields);
        _additionalFields.AddRange(decisionFields);

        // Create checkboxes for each field
        foreach (var field in personalFields)
        {
            var checkbox = new CheckBox
            {
                Content = field.DisplayName,
                Tag = field,
                Margin = new Thickness(0, 4),
                FontSize = 13
            };
            PersonalFieldsPanel.Children.Add(checkbox);
        }

        foreach (var field in applicationFields)
        {
            var checkbox = new CheckBox
            {
                Content = field.DisplayName,
                Tag = field,
                Margin = new Thickness(0, 4),
                FontSize = 13
            };
            ApplicationFieldsPanel.Children.Add(checkbox);
        }

        foreach (var field in decisionFields)
        {
            var checkbox = new CheckBox
            {
                Content = field.DisplayName,
                Tag = field,
                Margin = new Thickness(0, 4),
                FontSize = 13
            };
            DecisionFieldsPanel.Children.Add(checkbox);
        }
    }

    private async void InitializeDataAsync()
    {
        StatusLabel.Text = "Loading baked data...";
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
            // Get selected additional fields from checked checkboxes
            List<string> selectedFieldNames = new List<string>();

            // Get all checkboxes from all three panels
            var allCheckboxes = PersonalFieldsPanel.Children.OfType<CheckBox>()
                .Concat(ApplicationFieldsPanel.Children.OfType<CheckBox>())
                .Concat(DecisionFieldsPanel.Children.OfType<CheckBox>());

            foreach (var checkbox in allCheckboxes)
            {
                if (checkbox.IsChecked == true && checkbox.Tag is AdditionalField field)
                {
                    selectedFieldNames.Add(field.PropertyName);
                }
            }

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

                        var outputRecord = new OutputRecord {
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
                        };

                        // Populate additional fields using reflection
                        foreach (var propertyName in selectedFieldNames)
                        {
                            var property = typeof(ApplicationRecord).GetProperty(propertyName);
                            if (property != null)
                            {
                                var value = property.GetValue(app) as string;
                                outputRecord.AdditionalFieldValues[propertyName] = value;
                            }
                        }

                        outputRecords.Add(outputRecord);
                    }

                    await Dispatcher.UIThread.InvokeAsync(() => {
                        var item = ProcessingItems.FirstOrDefault(p => p.StudentNo == inTray.StudentNo);
                        if (item != null) { item.Status = app != null ? "✓ Done" : "⚠️ Missing"; item.UkGrade = classification; }
                        MainProgressBar.Value = (double)current / inTrayRecords.Count * 100;
                        FooterStatus.Text = $"{current}/{inTrayRecords.Count} Processed";
                    });
                }

                _csvService.GenerateOutputFiles(outputRecords, _outputFolderPath, selectedFieldNames.Count > 0 ? selectedFieldNames : null);

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

private void SetVersion()
{
    var v = Assembly.GetExecutingAssembly().GetName().Version;
    if (v != null)
    {
        // Major = 1, Minor = 0, Build = 25
        VersionLabel.Text = $"v{v.Major}.{v.Minor}.{v.Build}";
    }
}
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

            if (!File.Exists(soundPath))
            {
                // Portable dev path: look 3 levels up from bin/Debug/net10.0
                string devPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "audio", "confirmed.mp3");
                if (File.Exists(devPath)) soundPath = devPath;
            }

            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            {
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
                    string fullPath = Path.GetFullPath(soundPath);
                    mciSendString($"open \"{fullPath}\" type mpegvideo alias confirm", null, 0, IntPtr.Zero);
                    mciSendString("play confirm from 0", null, 0, IntPtr.Zero);
                }
            }
        }
        catch (Exception ex) { Debug.WriteLine($"Audio error: {ex.Message}"); }
    }

    private void LogStatus(string message) => Dispatcher.UIThread.Post(() => StatusLog.Text += $"[{DateTime.Now:HH:mm:ss}] {message}\n");
    private void ClearLogButton_Click(object? sender, RoutedEventArgs e) => StatusLog.Text = string.Empty;
    private void ExitButton_Click(object? sender, RoutedEventArgs e) => Environment.Exit(0);
    
    private void ResetButton_Click(object? sender, RoutedEventArgs e)
    {
        _inTrayFilePath = string.Empty;
        _appReportsFilePath = string.Empty;
        _outputFolderPath = string.Empty;

        InTrayFileLabel.Text = "No file selected";
        AppReportsFileLabel.Text = "No file selected";
        OutputFolderLabel.Text = "No folder selected";

        ProcessingItems.Clear();
        MainProgressBar.Value = 0;
        FooterStatus.Text = "Ready";
        StatusLog.Text = string.Empty;

        CheckReadyToProcess();
        LogStatus("Application reset.");
    }
}
