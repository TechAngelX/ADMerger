// MainForm.cs

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ADMerger.Configuration;
using ADMerger.Models;
using ADMerger.Services;
using ADMerger.UI.Controls;
using ADMerger.Utilities;
using OfficeOpenXml;

namespace ADMerger
{
    public partial class MainForm : Form
    {
        private readonly ICsvService _csvService;
        private readonly IEquivalencyService _equivalencyService;
        private readonly IRankingService _rankingService;
        private readonly IGradeClassificationService _gradeService;
        
        private string _document1Path = "";
        private string _document2Path = "";
        private List<InTrayRecord> _document1Data = new List<InTrayRecord>();
        private List<ApplicationRecord> _document2Data = new List<ApplicationRecord>();
        private string _lastOutputPath = "";
        
        private ModernFilePanel _doc1Panel;
        private ModernFilePanel _doc2Panel;
        private ModernButton _processButton;
        private ModernButton _exitButton;
        private ModernButton _openOutputButton;
        private RichTextBox _statusBox;
        
        public MainForm(
            ICsvService csvService,
            IEquivalencyService equivalencyService,
            IRankingService rankingService,
            IGradeClassificationService gradeService)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            _csvService = csvService ?? throw new ArgumentNullException(nameof(csvService));
            _equivalencyService = equivalencyService ?? throw new ArgumentNullException(nameof(equivalencyService));
            _rankingService = rankingService ?? throw new ArgumentNullException(nameof(rankingService));
            _gradeService = gradeService ?? throw new ArgumentNullException(nameof(gradeService));
            
            InitializeComponent();
            LoadData();
        }
        
        private void LoadData()
        {
            try
            {
                _equivalencyService.LoadEquivalencies();
                UpdateStatus($"Loaded {_equivalencyService.Count} country equivalencies");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Warning: Could not load equivalencies: {ex.Message}");
            }
            
            try
            {
                _rankingService.LoadRankings();
                UpdateStatus($"Loaded {_rankingService.Count} THE World University Rankings");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Warning: Could not load THE Rankings: {ex.Message}");
            }
        }
        
        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            this.Text = "AD Merger";
            this.ClientSize = new Size(900, 700);
            this.BackColor = Color.White;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;

            Panel headerPanel = CreateHeaderPanel();
            this.Controls.Add(headerPanel);

            int yPos = 120;
            yPos = CreateDocument1Section(yPos);
            yPos = CreateDocument2Section(yPos);

            _processButton = new ModernButton();
            _processButton.Text = "Process Files";
            _processButton.Location = new Point(30, yPos);
            _processButton.Size = new Size(180, 45);
            _processButton.Enabled = false;
            _processButton.Click += ProcessFiles_Click;
            _processButton.SetRounded();
            this.Controls.Add(_processButton);

            yPos += 65;
            yPos = CreateStatusSection(yPos);
            CreateBottomButtons(yPos);

            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private Panel CreateHeaderPanel()
        {
            Panel headerPanel = new Panel();
            headerPanel.Location = new Point(0, 0);
            headerPanel.Size = new Size(900, 90);
            headerPanel.BackColor = ColorTranslator.FromHtml("#3B82F6");
            headerPanel.Paint += (s, e) =>
            {
                using (LinearGradientBrush brush = new LinearGradientBrush(
                    headerPanel.ClientRectangle,
                    ColorTranslator.FromHtml("#3B82F6"),
                    ColorTranslator.FromHtml("#2563EB"),
                    LinearGradientMode.Horizontal))
                {
                    e.Graphics.FillRectangle(brush, headerPanel.ClientRectangle);
                }
            };

            Label titleLabel = new Label();
            titleLabel.Text = "AD Merger";
            titleLabel.Font = new Font("Segoe UI", 24F, FontStyle.Bold);
            titleLabel.ForeColor = Color.White;
            titleLabel.AutoSize = true;
            titleLabel.Location = new Point(30, 15);
            titleLabel.BackColor = Color.Transparent;
            headerPanel.Controls.Add(titleLabel);

            Label subtitleLabel = new Label();
            subtitleLabel.Text = "Admissions Data Merger";
            subtitleLabel.Font = new Font("Segoe UI", 11F, FontStyle.Regular);
            subtitleLabel.ForeColor = ColorTranslator.FromHtml("#DBEAFE");
            subtitleLabel.AutoSize = true;
            subtitleLabel.Location = new Point(33, 55);
            subtitleLabel.BackColor = Color.Transparent;
            headerPanel.Controls.Add(subtitleLabel);

            return headerPanel;
        }

        private int CreateDocument1Section(int yPos)
        {
            Label doc1Label = new Label();
            doc1Label.Text = "Document 1 (In-tray - New Applicants)";
            doc1Label.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            doc1Label.ForeColor = ColorTranslator.FromHtml("#334155");
            doc1Label.Location = new Point(30, yPos);
            doc1Label.AutoSize = true;
            this.Controls.Add(doc1Label);

            _doc1Panel = new ModernFilePanel("Click or drag CSV file for Document 1", 30, yPos + 30);
            _doc1Panel.Click += (s, e) => SelectDocument1();
            _doc1Panel.SetDropHandler(filePath => 
            {
                _document1Path = filePath;
                LoadInTrayData();
            });
            this.Controls.Add(_doc1Panel);

            return yPos + 140;
        }

        private int CreateDocument2Section(int yPos)
        {
            Label doc2Label = new Label();
            doc2Label.Text = "Document 2 (Application Reports)";
            doc2Label.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            doc2Label.ForeColor = ColorTranslator.FromHtml("#334155");
            doc2Label.Location = new Point(30, yPos);
            doc2Label.AutoSize = true;
            this.Controls.Add(doc2Label);

            _doc2Panel = new ModernFilePanel("Click or drag CSV file for Document 2", 30, yPos + 30);
            _doc2Panel.Click += (s, e) => SelectDocument2();
            _doc2Panel.SetDropHandler(filePath => 
            {
                if (_document1Data.Count == 0)
                {
                    MessageBox.Show("Please load Document 1 first.", "Info", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                _document2Path = filePath;
                LoadApplicationReports();
            });
            _doc2Panel.Enabled = false;
            this.Controls.Add(_doc2Panel);

            return yPos + 140;
        }

        private int CreateStatusSection(int yPos)
        {
            Label statusLabel = new Label();
            statusLabel.Text = "Status";
            statusLabel.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            statusLabel.ForeColor = ColorTranslator.FromHtml("#334155");
            statusLabel.Location = new Point(30, yPos);
            statusLabel.AutoSize = true;
            this.Controls.Add(statusLabel);

            _statusBox = new RichTextBox();
            _statusBox.Location = new Point(30, yPos + 25);
            _statusBox.Size = new Size(830, 100);
            _statusBox.ReadOnly = true;
            _statusBox.BackColor = ColorTranslator.FromHtml("#F8FAFC");
            _statusBox.BorderStyle = BorderStyle.FixedSingle;
            _statusBox.Font = new Font("Consolas", 9F);
            _statusBox.ForeColor = ColorTranslator.FromHtml("#475569");
            _statusBox.Text = "Ready. Click or drag CSV file for Document 1...";
            this.Controls.Add(_statusBox);

            return yPos + 135;
        }

        private void CreateBottomButtons(int yPos)
        {
            _exitButton = new ModernButton();
            _exitButton.Text = "Exit";
            _exitButton.Location = new Point(30, yPos);
            _exitButton.Size = new Size(120, 40);
            _exitButton.Click += (s, e) => Application.Exit();
            _exitButton.SetSecondary();
            this.Controls.Add(_exitButton);

            _openOutputButton = new ModernButton();
            _openOutputButton.Text = "Open Output Folder";
            _openOutputButton.Location = new Point(160, yPos);
            _openOutputButton.Size = new Size(180, 40);
            _openOutputButton.Enabled = false;
            _openOutputButton.Click += OpenOutputFolder_Click;
            this.Controls.Add(_openOutputButton);
        }

        private void UpdateStatus(string message)
        {
            if (_statusBox.InvokeRequired)
            {
                _statusBox.Invoke(new Action(() => UpdateStatus(message)));
                return;
            }
            _statusBox.AppendText(message + "\n");
            _statusBox.SelectionStart = _statusBox.Text.Length;
            _statusBox.ScrollToCaret();
        }

        private void SelectDocument1()
        {
            var dialog = new OpenFileDialog();
            dialog.Title = "Select Document 1 (Department In-tray - New Applicants CSV)";
            dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                _document1Path = dialog.FileName;
                LoadInTrayData();
            }
        }

        private void LoadInTrayData()
        {
            try
            {
                _document1Data = _csvService.LoadInTrayRecords(_document1Path);
                
                _doc1Panel.SetFileLoaded(Path.GetFileName(_document1Path), _document1Data.Count);
                UpdateStatus($"Document 1 loaded: {_document1Data.Count} new applicants");
                
                _doc2Panel.Enabled = true;
                _doc2Panel.UpdateText("Click or drag CSV file for Document 2");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Document 1: " + ex.Message, "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus($"ERROR: {ex.Message}");
            }
        }

        private void SelectDocument2()
        {
            if (_document1Data.Count == 0)
            {
                MessageBox.Show("Please load Document 1 first.", "Info", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var dialog = new OpenFileDialog();
            dialog.Title = "Select Document 2 (Department Application Reports CSV)";
            dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                _document2Path = dialog.FileName;
                LoadApplicationReports();
            }
        }

        private void LoadApplicationReports()
        {
            try
            {
                _document2Data = _csvService.LoadApplicationRecords(_document2Path);
                
                _doc2Panel.SetFileLoaded(Path.GetFileName(_document2Path), _document2Data.Count);
                UpdateStatus($"Document 2 loaded: {_document2Data.Count} application records");
                
                _processButton.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Document 2: " + ex.Message, "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus($"ERROR: {ex.Message}");
            }
        }

        private void ProcessFiles_Click(object sender, EventArgs e)
        {
            try
            {
                using (var folderDialog = new FolderBrowserDialog())
                {
                    folderDialog.Description = "Select folder to save output CSV files";
                    folderDialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    
                    if (folderDialog.ShowDialog() != DialogResult.OK)
                    {
                        UpdateStatus("Processing cancelled by user.");
                        return;
                    }
                    
                    _processButton.Enabled = false;
                    UpdateStatus("\nProcessing and cross-referencing data...");

                    var results = CrossReferenceData();
                    var outputPath = _csvService.GenerateOutputFiles(results, folderDialog.SelectedPath);
                    
                    _lastOutputPath = outputPath.Split('\n')[0];
                    _openOutputButton.Enabled = true;
                    
                    UpdateStatus($"\nSUCCESS! Matched {results.Count}/{_document1Data.Count} applicants");
                    UpdateStatus($"Output files created:\n{outputPath}");
                    
                    MessageBox.Show($"Processing complete!\n\nMatched {results.Count} out of {_document1Data.Count} new applicants.\n\nOutput files saved to:\n{outputPath}", 
                        "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error processing files: " + ex.Message, "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus($"\nERROR: {ex.Message}");
            }
            finally
            {
                _processButton.Enabled = true;
            }
        }

        private void OpenOutputFolder_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_lastOutputPath))
            {
                try
                {
                    string folder = Path.GetDirectoryName(_lastOutputPath);
                    System.Diagnostics.Process.Start("explorer.exe", folder);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Could not open folder: " + ex.Message, "Error", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private List<OutputRecord> CrossReferenceData()
        {
            var results = new List<OutputRecord>();
            
            foreach (var inTrayRecord in _document1Data)
            {
                var match = _document2Data.FirstOrDefault(app => app.ApplicantID == inTrayRecord.StudentNo);
                
                if (match != null)
                {
                    var programmeCode = ProgrammeMapping.GetCode(match.Programme);
                    var ukGrade = _gradeService.DetermineUKClassification(
                        match.OverallGradeGPA, 
                        match.EquivalencyNote, 
                        match.CountryOfStudy);
                    var theRanking = _rankingService.GetRanking(match.InstitutionName);
                    
                    results.Add(new OutputRecord
                    {
                        ReceivedDate = DateFormatter.FormatDate(inTrayRecord.ReceivedOn),
                        DueDate = DateFormatter.CalculateDueDate(inTrayRecord.ReceivedOn),
                        StudentNo = inTrayRecord.StudentNo,
                        Programme = programmeCode,
                        Forename = match.Forename,
                        Surname = match.Surname,
                        Gender = match.Gender,
                        DateOfBirth = DateFormatter.FormatDate(match.DateOfBirth),
                        FeeStatus = match.FeeStatus,
                        CountryOfStudy = match.CountryOfStudy,
                        CountryOfNationality = match.CountryOfNationality,
                        QualificationName = match.QualificationName,
                        DegreeSubject = match.DegreeSubject,
                        InstitutionName = match.InstitutionName,
                        THERanking = theRanking,
                        OverallGradeGPA = match.OverallGradeGPA,
                        EquivalencyNote = match.EquivalencyNote,
                        UKGrade = ukGrade
                    });
                }
            }
            
            return results;
        }
    }
}