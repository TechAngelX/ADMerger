using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;
using ADMerger.Models;
using ADMerger.Services;
using ADMerger.UI;

namespace ADMerger
{
    public partial class MainForm : Form
    {
        private readonly DataService dataService;
        private readonly RankingService rankingService;
        private readonly GradeClassificationService gradeService;
        private readonly OutputService outputService;
        
        private string document1Path = "";
        private string document2Path = "";
        private List<InTrayRecord> document1Data = new List<InTrayRecord>();
        private List<ApplicationRecord> document2Data = new List<ApplicationRecord>();
        private string lastOutputPath = "";
        
        private ModernFilePanel doc1Panel;
        private ModernFilePanel doc2Panel;
        private ModernButton processButton;
        private ModernButton exitButton;
        private ModernButton openOutputButton;
        private RichTextBox statusBox;
        
        public MainForm()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            dataService = new DataService();
            rankingService = new RankingService();
            
            var equivalencies = dataService.LoadDegreeEquivalencies();
            gradeService = new GradeClassificationService(equivalencies);
            outputService = new OutputService();
            
            InitializeComponent();
            
            rankingService.LoadTHERankings();
            UpdateStatus($"Loaded {equivalencies.Count} country equivalencies");
            UpdateStatus($"Loaded {rankingService.LoadedCount} THE World University Rankings");
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

            Label doc1Label = CreateLabel("Document 1 (In-tray - New Applicants)", yPos);
            this.Controls.Add(doc1Label);

            doc1Panel = new ModernFilePanel("Click or drag CSV file for Document 1", 30, yPos + 30);
            doc1Panel.Click += (s, e) => SelectDocument1();
            doc1Panel.SetDropHandler(filePath => 
            {
                document1Path = filePath;
                LoadInTrayData();
            });
            this.Controls.Add(doc1Panel);

            yPos += 140;

            Label doc2Label = CreateLabel("Document 2 (Application Reports)", yPos);
            this.Controls.Add(doc2Label);

            doc2Panel = new ModernFilePanel("Click or drag CSV file for Document 2", 30, yPos + 30);
            doc2Panel.Click += (s, e) => SelectDocument2();
            doc2Panel.SetDropHandler(filePath => 
            {
                if (document1Data.Count == 0)
                {
                    MessageBox.Show("Please load Document 1 first.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                document2Path = filePath;
                LoadApplicationReports();
            });
            doc2Panel.Enabled = false;
            this.Controls.Add(doc2Panel);

            yPos += 140;

            processButton = new ModernButton();
            processButton.Text = "Process Files";
            processButton.Location = new Point(30, yPos);
            processButton.Size = new Size(180, 45);
            processButton.Enabled = false;
            processButton.Click += ProcessFiles_Click;
            processButton.SetRounded();
            this.Controls.Add(processButton);

            yPos += 65;

            Label statusLabel = CreateLabel("Status", yPos);
            this.Controls.Add(statusLabel);

            statusBox = new RichTextBox();
            statusBox.Location = new Point(30, yPos + 25);
            statusBox.Size = new Size(830, 100);
            statusBox.ReadOnly = true;
            statusBox.BackColor = ColorTranslator.FromHtml("#F8FAFC");
            statusBox.BorderStyle = BorderStyle.FixedSingle;
            statusBox.Font = new Font("Consolas", 9F);
            statusBox.ForeColor = ColorTranslator.FromHtml("#475569");
            statusBox.Text = "Ready. Click or drag CSV file for Document 1...";
            this.Controls.Add(statusBox);

            yPos += 135;

            exitButton = new ModernButton();
            exitButton.Text = "Exit";
            exitButton.Location = new Point(30, yPos);
            exitButton.Size = new Size(120, 40);
            exitButton.Click += (s, e) => Application.Exit();
            exitButton.SetSecondary();
            this.Controls.Add(exitButton);

            openOutputButton = new ModernButton();
            openOutputButton.Text = "Open Output Folder";
            openOutputButton.Location = new Point(160, yPos);
            openOutputButton.Size = new Size(180, 40);
            openOutputButton.Enabled = false;
            openOutputButton.Click += OpenOutputFolder_Click;
            this.Controls.Add(openOutputButton);

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

        private Label CreateLabel(string text, int yPos)
        {
            Label label = new Label();
            label.Text = text;
            label.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            label.ForeColor = ColorTranslator.FromHtml("#334155");
            label.Location = new Point(30, yPos);
            label.AutoSize = true;
            return label;
        }

        private void UpdateStatus(string message)
        {
            if (statusBox.InvokeRequired)
            {
                statusBox.Invoke(new Action(() => UpdateStatus(message)));
                return;
            }
            statusBox.AppendText(message + "\n");
            statusBox.SelectionStart = statusBox.Text.Length;
            statusBox.ScrollToCaret();
        }

        private void SelectDocument1()
        {
            var dialog = new OpenFileDialog();
            dialog.Title = "Select Document 1 (Department In-tray - New Applicants CSV)";
            dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                document1Path = dialog.FileName;
                LoadInTrayData();
            }
        }

        private void LoadInTrayData()
        {
            try
            {
                document1Data = dataService.LoadInTrayData(document1Path);
                
                doc1Panel.SetFileLoaded(System.IO.Path.GetFileName(document1Path), document1Data.Count);
                UpdateStatus($"Document 1 loaded: {document1Data.Count} new applicants");
                
                doc2Panel.Enabled = true;
                doc2Panel.UpdateText("Click or drag CSV file for Document 2");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Document 1: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus($"ERROR: {ex.Message}");
            }
        }

        private void SelectDocument2()
        {
            if (document1Data.Count == 0)
            {
                MessageBox.Show("Please load Document 1 first.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var dialog = new OpenFileDialog();
            dialog.Title = "Select Document 2 (Department Application Reports CSV)";
            dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                document2Path = dialog.FileName;
                LoadApplicationReports();
            }
        }

        private void LoadApplicationReports()
        {
            try
            {
                document2Data = dataService.LoadApplicationData(document2Path);
                
                doc2Panel.SetFileLoaded(System.IO.Path.GetFileName(document2Path), document2Data.Count);
                UpdateStatus($"Document 2 loaded: {document2Data.Count} application records");
                
                processButton.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Document 2: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    
                    processButton.Enabled = false;
                    UpdateStatus("\nProcessing and cross-referencing data...");

                    var results = CrossReferenceData();
                    var outputPath = outputService.GenerateOutputFiles(results, folderDialog.SelectedPath);
                    
                    lastOutputPath = outputPath.Split('\n')[0];
                    openOutputButton.Enabled = true;
                    
                    UpdateStatus($"\nSUCCESS! Matched {results.Count}/{document1Data.Count} applicants");
                    UpdateStatus($"Output files created:\n{outputPath}");
                    
                    MessageBox.Show($"Processing complete!\n\nMatched {results.Count} out of {document1Data.Count} new applicants.\n\nOutput files saved to:\n{outputPath}", 
                        "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error processing files: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus($"\nERROR: {ex.Message}");
            }
            finally
            {
                processButton.Enabled = true;
            }
        }

        private List<OutputRecord> CrossReferenceData()
        {
            var results = new List<OutputRecord>();
            
            foreach (var inTrayRecord in document1Data)
            {
                var match = document2Data.FirstOrDefault(app => app.ApplicantID == inTrayRecord.StudentNo);
                
                if (match != null)
                {
                    var programmeCode = outputService.GetProgrammeCode(match.Programme);
                    var ukGrade = gradeService.DetermineUKClassification(match.OverallGradeGPA, match.EquivalencyNote, match.CountryOfStudy);
                    var theRanking = rankingService.GetTHERanking(match.InstitutionName);
                    
                    results.Add(new OutputRecord
                    {
                        ReceivedDate = outputService.FormatDate(inTrayRecord.ReceivedOn),
                        DueDate = outputService.CalculateDueDate(inTrayRecord.ReceivedOn),
                        StudentNo = inTrayRecord.StudentNo,
                        Programme = programmeCode,
                        Forename = match.Forename,
                        Surname = match.Surname,
                        Gender = match.Gender,
                        DateOfBirth = outputService.FormatDate(match.DateOfBirth),
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

        private void OpenOutputFolder_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(lastOutputPath))
            {
                try
                {
                    string folder = System.IO.Path.GetDirectoryName(lastOutputPath);
                    System.Diagnostics.Process.Start("explorer.exe", folder);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Could not open folder: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
