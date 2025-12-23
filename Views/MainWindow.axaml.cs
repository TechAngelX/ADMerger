// Views/MainWindow.axaml.cs

using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using ADMerger.Services;
using ADMerger.Models;
using ADMerger.Configuration;  
using ADMerger.Utilities;       
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ADMerger.Views;

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
    
    public MainWindow()
    {
        InitializeComponent();
        
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
        var infoVersion = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>();
        var version = infoVersion?.InformationalVersion ?? "1.0.0";
        VersionLabel.Text = $"UCL Admissions Data Processor v{version}";
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
                    AppleUniformTypeIdentifiers = new[] 
                    { 
                        "public.comma-separated-values-text", 
                        "org.openxmlformats.spreadsheetml.sheet" 
                    },
                    MimeTypes = new[] 
                    { 
                        "text/csv", 
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" 
                    }
                }
            }
        });
        
        if (files.Count > 0)
        {
            _inTrayFilePath = files[0].Path.LocalPath;
            InTrayFileLabel.Text = Path.GetFileName(_inTrayFilePath);
            LogStatus($"Selected InTray file: {Path.GetFileName(_inTrayFilePath)}");
            CheckReadyToProcess();
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
                    AppleUniformTypeIdentifiers = new[] 
                    { 
                        "public.comma-separated-values-text", 
                        "org.openxmlformats.spreadsheetml.sheet" 
                    },
                    MimeTypes = new[] 
                    { 
                        "text/csv", 
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" 
                    }
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
            StatusLabel.Text = "Ready to process! Click the button above.";
            FooterStatus.Text = "Ready to process";
        }
        else
        {
            StatusLabel.Text = "Load all files to begin";
            FooterStatus.Text = "Waiting for files...";
        }
    }
    
    private async void ProcessButton_Click(object? sender, RoutedEventArgs e)
    {
        ProcessButton.IsEnabled = false;
        BrowseInTrayButton.IsEnabled = false;
        BrowseAppReportsButton.IsEnabled = false;
        BrowseOutputButton.IsEnabled = false;
        
        try
        {
            LogStatus("=== Starting Processing ===");
            
            LogStatus("Loading InTray records...");
            var inTrayRecords = _csvService.LoadInTrayRecords(_inTrayFilePath);
            LogStatus($"Loaded {inTrayRecords.Count} InTray records");
            
            LogStatus("Loading Application records...");
            var appRecords = _csvService.LoadApplicationRecords(_appReportsFilePath);
            LogStatus($"Loaded {appRecords.Count} Application records");
            
            LogStatus("Merging and enriching data...");
            var outputRecords = MergeAndEnrichData(inTrayRecords, appRecords);
            LogStatus($"Generated {outputRecords.Count} output records");
            
            LogStatus("Generating Excel files...");
            var outputPaths = _csvService.GenerateOutputFiles(outputRecords, _outputFolderPath);
            
            LogStatus("=== Processing Complete ===");
            LogStatus("Generated files:");
            foreach (var path in outputPaths)
            {
                LogStatus($"  - {Path.GetFileName(path)}");
            }
            
            FooterStatus.Text = $"Complete! Generated {outputPaths.Count} files";
            
            await ShowMessageBoxAsync("Success", 
                $"Processing complete!\n\nGenerated {outputPaths.Count} Excel files in:\n{_outputFolderPath}");
        }
        catch (Exception ex)
        {
            LogStatus($"ERROR: {ex.Message}");
            FooterStatus.Text = "Error occurred - check log";
            
            await ShowMessageBoxAsync("Error", $"Processing failed:\n\n{ex.Message}");
        }
        finally
        {
            ProcessButton.IsEnabled = true;
            BrowseInTrayButton.IsEnabled = true;
            BrowseAppReportsButton.IsEnabled = true;
            BrowseOutputButton.IsEnabled = true;
        }
    }
    
    private List<OutputRecord> MergeAndEnrichData(
        List<InTrayRecord> inTrayRecords, 
        List<ApplicationRecord> appRecords)
    {
        var outputRecords = new List<OutputRecord>();
        
        foreach (var inTray in inTrayRecords)
        {
            var app = appRecords.FirstOrDefault(a => a.ApplicantID == inTray.StudentNo);
            
            if (app == null)
            {
                LogStatus($"Warning: No application record found for {inTray.StudentNo}");
                continue;
            }
            
            var programmeCode = ProgrammeMapping.GetCode(app.Programme);
            
            var ukGrade = _gradeService.DetermineUKClassification(
                app.OverallGradeGPA,
                app.EquivalencyNote,
                app.CountryOfStudy
            );
            
            var theRanking = _rankingService.GetRanking(app.InstitutionName ?? "");
            
            var output = new OutputRecord
            {
                ReceivedDate = DateFormatter.FormatDate(inTray.ReceivedOn),
                DueDate = DateFormatter.CalculateDueDate(inTray.ReceivedOn),
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
            };
            
            outputRecords.Add(output);
        }
        
        return outputRecords;
    }
    
    private void ClearLogButton_Click(object? sender, RoutedEventArgs e)
    {
        StatusLog.Text = string.Empty;
        LogStatus("Log cleared");
    }
    
    private void ResetButton_Click(object? sender, RoutedEventArgs e)
    {
        _inTrayFilePath = string.Empty;
        _appReportsFilePath = string.Empty;
        _outputFolderPath = string.Empty;
        
        InTrayFileLabel.Text = "No file selected";
        AppReportsFileLabel.Text = "No file selected";
        OutputFolderLabel.Text = "No folder selected";
        StatusLog.Text = string.Empty;
        
        CheckReadyToProcess();
        LogStatus("Application reset");
    }
    
    private void ExitButton_Click(object? sender, RoutedEventArgs e)
    {
        Close();
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
            Content = new StackPanel
            {
                Margin = new Avalonia.Thickness(20),
                Spacing = 15,
                Children =
                {
                    new TextBlock 
                    { 
                        Text = message, 
                        TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                        FontSize = 14
                    },
                    new Button
                    {
                        Content = "OK",
                        HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
                        Padding = new Avalonia.Thickness(40, 12),
                        Background = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#4299E1")),
                        Foreground = Avalonia.Media.Brushes.White,
                        CornerRadius = new Avalonia.CornerRadius(4)
                    }
                }
            }
        };
        
        ((Button)((StackPanel)messageBox.Content).Children[1]).Click += (s, e) => messageBox.Close();
        
        await messageBox.ShowDialog(this);
    }
}
