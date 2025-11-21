// ./MainForm.cs

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using CsvHelper;
using System.Globalization;
using CsvHelper.Configuration;

namespace ADMerger
{
    public partial class MainForm : Form
    {
        private string document1Path = "";
        private string document2Path = "";
        private List<InTrayRecord> document1Data = new List<InTrayRecord>();
        private List<ApplicationRecord> document2Data = new List<ApplicationRecord>();
        
        private ModernFilePanel doc1Panel;
        private ModernFilePanel doc2Panel;
        private ModernButton processButton;
        private ModernButton exitButton;
        private ModernButton openOutputButton;
        private RichTextBox statusBox;
        
        private Dictionary<string, DegreeEquivalency> degreeEquivalencies = new Dictionary<string, DegreeEquivalency>();
        private string lastOutputPath = "";
        
        public MainForm()
        {
            InitializeComponent();
            LoadDegreeEquivalencies();
        }
        
        private void LoadDegreeEquivalencies()
        {
            try
            {
                StreamReader reader = null;
                
                // Try to load from file system first (for dev/debug)
                string csvPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "ucl_degree_equivalencies_FINAL.csv");
                
                if (File.Exists(csvPath))
                {
                    // Dev mode: load from file
                    reader = new StreamReader(csvPath);
                    UpdateStatus("Loading equivalencies from data folder (dev mode)...");
                }
                else
                {
                    // Release mode: load from embedded resource
                    var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                    var resourceName = "ADMerger.data.ucl_degree_equivalencies_FINAL.csv";
                    var stream = assembly.GetManifestResourceStream(resourceName);
                    
                    if (stream == null)
                    {
                        UpdateStatus("Warning: Embedded equivalencies data not found. UK grade mapping will use basic rules only.");
                        return;
                    }
                    
                    reader = new StreamReader(stream);
                    UpdateStatus("Loading equivalencies from embedded resource (release mode)...");
                }
                
                using (reader)
                {
                    var config = new CsvConfiguration(CultureInfo.InvariantCulture);
                    config.HeaderValidated = null;
                    config.MissingFieldFound = null;
                    config.Delimiter = "\t";
                    
                    using var csv = new CsvReader(reader, config);
                    
                    csv.Read();
                    csv.ReadHeader();
                    
                    while (csv.Read())
                    {
                        try
                        {
                            string country = csv.GetField(0)?.Trim().TrimStart('\'');
                            string third = csv.GetField(1)?.Trim().TrimStart('\'').TrimStart('<');
                            string secondLower = csv.GetField(2)?.Trim().TrimStart('\'');
                            string secondUpper = csv.GetField(3)?.Trim().TrimStart('\'');
                            string first = csv.GetField(4)?.Trim().TrimStart('\'');
                            
                            if (!string.IsNullOrWhiteSpace(country))
                            {
                                degreeEquivalencies[country] = new DegreeEquivalency
                                {
                                    Country = country,
                                    Third = third,
                                    SecondLower = secondLower,
                                    SecondUpper = secondUpper,
                                    First = first
                                };
                            }
                        }
                        catch { }
                    }
                    
                    UpdateStatus($"Loaded {degreeEquivalencies.Count} country equivalencies");
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"Warning: Could not load equivalencies: {ex.Message}");
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

            this.Controls.Add(headerPanel);

            int yPos = 120;

            Label doc1Label = new Label();
            doc1Label.Text = "Document 1 (In-tray - New Applicants)";
            doc1Label.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            doc1Label.ForeColor = ColorTranslator.FromHtml("#334155");
            doc1Label.Location = new Point(30, yPos);
            doc1Label.AutoSize = true;
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

            Label doc2Label = new Label();
            doc2Label.Text = "Document 2 (Application Reports)";
            doc2Label.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            doc2Label.ForeColor = ColorTranslator.FromHtml("#334155");
            doc2Label.Location = new Point(30, yPos);
            doc2Label.AutoSize = true;
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

            Label statusLabel = new Label();
            statusLabel.Text = "Status";
            statusLabel.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            statusLabel.ForeColor = ColorTranslator.FromHtml("#334155");
            statusLabel.Location = new Point(30, yPos);
            statusLabel.AutoSize = true;
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
                var config = new CsvConfiguration(CultureInfo.InvariantCulture);
                config.HeaderValidated = null;
                config.MissingFieldFound = null;
                
                using var reader = new StringReader(File.ReadAllText(document1Path));
                using var csv = new CsvReader(reader, config);
                document1Data = csv.GetRecords<InTrayRecord>().ToList();
                
                doc1Panel.SetFileLoaded(Path.GetFileName(document1Path), document1Data.Count);
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
                var config = new CsvConfiguration(CultureInfo.InvariantCulture);
                config.HeaderValidated = null;
                config.MissingFieldFound = null;
                
                using var reader = new StringReader(File.ReadAllText(document2Path));
                using var csv = new CsvReader(reader, config);
                document2Data = csv.GetRecords<ApplicationRecord>().ToList();
                
                doc2Panel.SetFileLoaded(Path.GetFileName(document2Path), document2Data.Count);
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
                // Ask user where to save output files
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
                    var outputPath = GenerateOutputFile(results, folderDialog.SelectedPath);
                    
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

        private void OpenOutputFolder_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(lastOutputPath))
            {
                try
                {
                    string folder = Path.GetDirectoryName(lastOutputPath);
                    System.Diagnostics.Process.Start("explorer.exe", folder);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Could not open folder: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private double? ParseGradeValue(string gradeStr)
        {
            if (string.IsNullOrWhiteSpace(gradeStr)) return null;
            
            gradeStr = gradeStr.TrimStart('\'').Trim();
            
            bool hadPercent = gradeStr.Contains("%");
            gradeStr = gradeStr.Replace("%", "").Trim();
            
            if (gradeStr.Contains("/"))
            {
                var parts = gradeStr.Split('/');
                if (parts.Length == 2 && 
                    double.TryParse(parts[0].Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out double numerator) &&
                    double.TryParse(parts[1].Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out double denominator) &&
                    denominator != 0)
                {
                    if (denominator == 100.0)
                    {
                        return numerator;
                    }
                    else
                    {
                        return (numerator / denominator) * 100.0;
                    }
                }
            }
            
            if (double.TryParse(gradeStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double result))
            {
                return result;
            }
            
            return null;
        }

        private string ParseUKGradeText(string gradeText)
        {
            if (string.IsNullOrWhiteSpace(gradeText)) return "??";
            
            string lower = gradeText.ToLower();
            
            if (lower.Contains("first") || lower.Contains("1st") || lower == "1.0" || lower == "1")
                return "1.0";
            
            if (lower.Contains("2.1") || lower.Contains("2:1") || lower.Contains("upper second"))
                return "2.1";
            
            if (lower.Contains("2.2") || lower.Contains("2:2") || lower.Contains("lower second"))
                return "2.2";
            
            if (lower.Contains("third") || lower.Contains("3rd") || lower == "3.0" || lower == "3")
                return "3.0";
            
            return "??";
        }

        private string FormatDate(string dateStr)
        {
            if (string.IsNullOrWhiteSpace(dateStr)) return "";
            
            string[] formats = {
                "dd-MMM-yy", "dd-MMM-yyyy",
                "dd/MM/yy", "dd/MM/yyyy",
                "yyyy-MM-dd", "MM/dd/yyyy"
            };
            
            foreach (var format in formats)
            {
                if (DateTime.TryParseExact(dateStr, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date))
                {
                    return date.ToString("dd/MM/yy");
                }
            }
            
            if (DateTime.TryParse(dateStr, out DateTime generalDate))
            {
                return generalDate.ToString("dd/MM/yy");
            }
            
            return dateStr;
        }

        private string CalculateDueDate(string receivedDateStr)
        {
            if (string.IsNullOrWhiteSpace(receivedDateStr)) return "";
            
            string[] formats = {
                "dd-MMM-yy", "dd-MMM-yyyy",
                "dd/MM/yy", "dd/MM/yyyy",
                "yyyy-MM-dd", "MM/dd/yyyy"
            };
            
            foreach (var format in formats)
            {
                if (DateTime.TryParseExact(receivedDateStr, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date))
                {
                    return date.AddDays(42).ToString("dd/MM/yy");
                }
            }
            
            if (DateTime.TryParse(receivedDateStr, out DateTime generalDate))
            {
                return generalDate.AddDays(42).ToString("dd/MM/yy");
            }
            
            return "";
        }

        private Dictionary<string, double> ParseThresholdsFromEquivalencyNote(string equivalencyNote)
        {
            var thresholds = new Dictionary<string, double>();
            
            var regex = new System.Text.RegularExpressions.Regex(@"(2\.2|2\.1|1st):\s*Bachelors\s*@\s*([0-9.]+)%");
            var matches = regex.Matches(equivalencyNote);
            
            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                string grade = match.Groups[1].Value;
                if (double.TryParse(match.Groups[2].Value, out double threshold))
                {
                    thresholds[grade] = threshold;
                }
            }
            
            return thresholds;
        }

        private string DetermineUKClassification(string overallGradeGPA, string equivalencyNote, string countryOfStudy)
        {
            if (string.IsNullOrWhiteSpace(countryOfStudy))
                return "??";
            
            if (countryOfStudy.ToLower().Contains("united kingdom"))
            {
                if (!string.IsNullOrWhiteSpace(overallGradeGPA))
                {
                    string grade = ParseUKGradeText(overallGradeGPA);
                    if (grade != "??") return grade;
                }
                
                return DetermineUKClassificationFromNote(equivalencyNote);
            }
            
            if (!string.IsNullOrWhiteSpace(equivalencyNote))
            {
                var noteLower = equivalencyNote.ToLower();
                if (noteLower.Contains("liverpool") || noteLower.Contains("reading"))
                {
                    var percentMatches = System.Text.RegularExpressions.Regex.Matches(equivalencyNote, @"(\d+(?:\.\d+)?)%");
                    double bestGrade = 0;
                    
                    foreach (System.Text.RegularExpressions.Match match in percentMatches)
                    {
                        if (double.TryParse(match.Groups[1].Value, out double grade))
                        {
                            if (grade > bestGrade && grade <= 100 && grade >= 30)
                            {
                                bestGrade = grade;
                            }
                        }
                    }
                    
                    if (bestGrade > 0)
                    {
                        if (bestGrade >= 70) return "1.0";
                        if (bestGrade >= 60) return "2.1";
                        if (bestGrade >= 50) return "2.2";
                        if (bestGrade >= 40) return "3.0";
                    }
                }
            }
            
            double? studentGrade = ParseGradeValue(overallGradeGPA);
            
            if (studentGrade == null && !string.IsNullOrWhiteSpace(equivalencyNote))
            {
                var gradeMatches = System.Text.RegularExpressions.Regex.Matches(
                    equivalencyNote, 
                    @"(?:grade of|average of|GPA of)\s*([0-9.]+(?:/[0-9.]+)?%?)"
                );
                
                if (gradeMatches.Count > 0)
                {
                    string gradeText = gradeMatches[gradeMatches.Count - 1].Groups[1].Value;
                    studentGrade = ParseGradeValue(gradeText);
                }
            }
            
            if (studentGrade != null && !string.IsNullOrWhiteSpace(equivalencyNote))
            {
                var thresholdsFromNote = ParseThresholdsFromEquivalencyNote(equivalencyNote);
                
                if (thresholdsFromNote.Count > 0)
                {
                    if (thresholdsFromNote.ContainsKey("1st") && studentGrade >= thresholdsFromNote["1st"])
                        return "1.0";
                    
                    if (thresholdsFromNote.ContainsKey("2.1") && studentGrade >= thresholdsFromNote["2.1"])
                        return "2.1";
                    
                    if (thresholdsFromNote.ContainsKey("2.2") && studentGrade >= thresholdsFromNote["2.2"])
                        return "2.2";
                    
                    return "3.0";
                }
            }
            
            string cleanCountry = countryOfStudy.Trim();
            
            if (studentGrade != null && degreeEquivalencies.ContainsKey(cleanCountry))
            {
                var equiv = degreeEquivalencies[cleanCountry];
                
                double? firstThreshold = ParseGradeValue(equiv.First);
                double? upperSecondThreshold = ParseGradeValue(equiv.SecondUpper);
                double? lowerSecondThreshold = ParseGradeValue(equiv.SecondLower);
                
                if (firstThreshold != null && studentGrade >= firstThreshold)
                    return "1.0";
                
                if (upperSecondThreshold != null && studentGrade >= upperSecondThreshold)
                    return "2.1";
                
                if (lowerSecondThreshold != null && studentGrade >= lowerSecondThreshold)
                    return "2.2";
                
                return "3.0";
            }
            
            return DetermineUKClassificationFromNote(equivalencyNote);
        }

        private string DetermineUKClassificationFromNote(string equivalencyNote)
        {
            if (string.IsNullOrWhiteSpace(equivalencyNote))
                return "??";
            
            var note = equivalencyNote.ToLower();
            
            var percentMatches = System.Text.RegularExpressions.Regex.Matches(equivalencyNote, @"(\d+(?:\.\d+)?)%");
            double bestGrade = 0;
            
            foreach (System.Text.RegularExpressions.Match match in percentMatches)
            {
                if (double.TryParse(match.Groups[1].Value, out double grade))
                {
                    if (grade > bestGrade && grade <= 100 && grade >= 30)
                    {
                        bestGrade = grade;
                    }
                }
            }
            
            if (bestGrade > 0)
            {
                if (bestGrade >= 70) return "1.0";
                if (bestGrade >= 60) return "2.1";
                if (bestGrade >= 50) return "2.2";
                if (bestGrade >= 40) return "3.0";
            }
            
            if (note.Contains("overall classification was 1st") || note.Contains("hons 1st") || note.Contains("first class")) return "1.0";
            if (note.Contains("overall classification was 2.1") || note.Contains("classification was 2.1") || note.Contains("upper second")) return "2.1";
            if (note.Contains("overall classification was 2.2") || note.Contains("classification was 2.2") || note.Contains("lower second")) return "2.2";
            if (note.Contains("overall classification was 3rd") || note.Contains("classification was 3rd") || note.Contains("third class")) return "3.0";
            
            return "??";
        }

        private List<OutputRecord> CrossReferenceData()
        {
            var results = new List<OutputRecord>();
            
            var programmeMapping = new Dictionary<string, string>
            {
                {"MSc Artificial Intelligence for Biomedicine and Healthcare", "AIBH"},
                {"MSc Artificial Intelligence for Sustainable Development", "AISD"},
                {"MSc Artificial Intelligence and Data Engineering", "AIDE"},
                {"MSc Information Security", "ISEC"},
                {"MSc Computational Finance", "CF"},
                {"MSc Financial Risk Management", "FRM"},
                {"MSc Financial Technology", "FT"},
                {"MSc Emerging Digital Technologies", "EDT"},
                {"MSc Machine Learning", "ML"},
                {"MSc Data Science and Machine Learning", "DSML"},
                {"MSc Computational Statistics and Machine Learning", "CSML"},
                {"MSc Robotics and Artificial Intelligence", "RAI"},
                {"MSc Systems Engineering for the Internet of Things", "SEIOT"},
                {"MSc Disability, Design and Innovation", "DDI"},
                {"MSc Computer Science", "CS"},
                {"MSc Software Systems Engineering", "SSE"},
                {"MSc Computer Graphics, Vision and Imaging", "CGVI"}
            };
            
            foreach (var inTrayRecord in document1Data)
            {
                var match = document2Data.FirstOrDefault(app => app.ApplicantID == inTrayRecord.StudentNo);
                
                if (match != null)
                {
                    var programmeCode = programmeMapping.ContainsKey(match.Programme) ? programmeMapping[match.Programme] : match.Programme;
                    var ukGrade = DetermineUKClassification(match.OverallGradeGPA, match.EquivalencyNote, match.CountryOfStudy);
                    
                    results.Add(new OutputRecord
                    {
                        ReceivedDate = FormatDate(inTrayRecord.ReceivedOn),
                        DueDate = CalculateDueDate(inTrayRecord.ReceivedOn),
                        StudentNo = inTrayRecord.StudentNo,
                        Programme = programmeCode,
                        Forename = match.Forename,
                        Surname = match.Surname,
                        Gender = match.Gender,
                        DateOfBirth = FormatDate(match.DateOfBirth),
                        FeeStatus = match.FeeStatus,
                        CountryOfStudy = match.CountryOfStudy,
                        CountryOfNationality = match.CountryOfNationality,
                        QualificationName = match.QualificationName,
                        DegreeSubject = match.DegreeSubject,
                        InstitutionName = match.InstitutionName,
                        OverallGradeGPA = match.OverallGradeGPA,
                        EquivalencyNote = match.EquivalencyNote,
                        UKGrade = ukGrade
                    });
                }
            }
            
            return results;
        }

        private string GenerateOutputFile(List<OutputRecord> data, string outputFolderPath)
        {
            var programmeGroups = data.GroupBy(record => record.Programme).ToList();
            var outputPaths = new List<string>();
            
            var columnOrder = new List<string>
            {
                "ReceivedDate",
                "DueDate",
                "StudentNo",
                "Programme",
                "Forename",
                "Surname",
                "Gender",
                "DateOfBirth",
                "FeeStatus",
                "CountryOfNationality",
                "QualificationName",
                "DegreeSubject",
                "InstitutionName",
                "CountryOfStudy",
                "EquivalencyNote",
                "OverallGradeGPA",
                "UKGrade",
                "Decision",
                "AT",
                "Note",
                "Progr. Adm",
                "Comment"
            };
            
            foreach (var group in programmeGroups)
            {
                var programme = group.Key;
                var records = group.ToList();
                
                var outputPath = Path.Combine(
                    outputFolderPath, 
                    programme + "_Latest_" + DateTime.Now.ToString("dd_MMM_yyyy_HHmm") + ".csv");
                
                using var writer = new StreamWriter(outputPath);
                
                writer.WriteLine(string.Join(",", columnOrder));
                
                foreach (var record in records)
                {
                    var values = new List<string>();
                    
                    foreach (var column in columnOrder)
                    {
                        string value = column switch
                        {
                            "ReceivedDate" => record.ReceivedDate ?? "",
                            "DueDate" => record.DueDate ?? "",
                            "StudentNo" => record.StudentNo ?? "",
                            "Programme" => record.Programme ?? "",
                            "Forename" => record.Forename ?? "",
                            "Surname" => record.Surname ?? "",
                            "Gender" => record.Gender ?? "",
                            "DateOfBirth" => record.DateOfBirth ?? "",
                            "FeeStatus" => record.FeeStatus ?? "",
                            "CountryOfNationality" => record.CountryOfNationality ?? "",
                            "QualificationName" => record.QualificationName ?? "",
                            "DegreeSubject" => record.DegreeSubject ?? "",
                            "InstitutionName" => record.InstitutionName ?? "",
                            "CountryOfStudy" => record.CountryOfStudy ?? "",
                            "EquivalencyNote" => record.EquivalencyNote ?? "",
                            "OverallGradeGPA" => record.OverallGradeGPA ?? "",
                            "UKGrade" => record.UKGrade ?? "",
                            _ => ""
                        };
                        
                        if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
                        {
                            value = "\"" + value.Replace("\"", "\"\"") + "\"";
                        }
                        
                        values.Add(value);
                    }
                    
                    writer.WriteLine(string.Join(",", values));
                }
                
                outputPaths.Add(outputPath);
            }
            
            return string.Join("\n", outputPaths);
        }
    }

    public class DegreeEquivalency
    {
        public string Country { get; set; }
        public string Third { get; set; }
        public string SecondLower { get; set; }
        public string SecondUpper { get; set; }
        public string First { get; set; }
    }

    public class ModernFilePanel : Panel
    {
        private Label label;
        private string originalText;
        private bool fileLoaded = false;
        private Action<string> onFileDropped;

        public ModernFilePanel(string text, int xPos, int yPos)
        {
            this.originalText = text;
            this.Location = new Point(xPos, yPos);
            this.Size = new Size(830, 90);
            this.BackColor = ColorTranslator.FromHtml("#F8FAFC");
            this.BorderStyle = BorderStyle.None;
            this.Cursor = Cursors.Hand;
            this.AllowDrop = true;
            
            this.Paint += (s, e) =>
            {
                using (Pen pen = new Pen(ColorTranslator.FromHtml("#CBD5E1"), 2))
                {
                    pen.DashStyle = DashStyle.Dash;
                    e.Graphics.DrawRectangle(pen, 1, 1, this.Width - 3, this.Height - 3);
                }
            };

            label = new Label();
            label.Text = text;
            label.Font = new Font("Segoe UI", 11F);
            label.ForeColor = ColorTranslator.FromHtml("#64748B");
            label.TextAlign = ContentAlignment.MiddleCenter;
            label.Dock = DockStyle.Fill;
            label.Cursor = Cursors.Hand;
            label.BackColor = Color.Transparent;
            label.Click += (s, e) => this.OnClick(e);
            
            label.MouseDown += (s, e) => this.OnMouseDown(e);
            label.MouseMove += (s, e) => this.OnMouseMove(e);
            label.MouseUp += (s, e) => this.OnMouseUp(e);
            
            this.Controls.Add(label);

            this.DragEnter += ModernFilePanel_DragEnter;
            this.DragDrop += ModernFilePanel_DragDrop;
            this.DragLeave += ModernFilePanel_DragLeave;
            this.DragOver += ModernFilePanel_DragOver;

            this.MouseEnter += (s, e) =>
            {
                if (this.Enabled && !fileLoaded)
                {
                    this.BackColor = ColorTranslator.FromHtml("#DBEAFE");
                    this.Invalidate();
                }
            };
            this.MouseLeave += (s, e) =>
            {
                if (!fileLoaded && this.Enabled)
                {
                    this.BackColor = ColorTranslator.FromHtml("#F8FAFC");
                    this.Invalidate();
                }
            };
        }

        private void ModernFilePanel_DragOver(object sender, DragEventArgs e)
        {
            if (!this.Enabled) return;
            
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length == 1 && files[0].EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    e.Effect = DragDropEffects.Copy;
                }
                else
                {
                    e.Effect = DragDropEffects.None;
                }
            }
        }

        private void ModernFilePanel_DragEnter(object sender, DragEventArgs e)
        {
            if (!this.Enabled) return;
            
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length == 1 && files[0].EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    e.Effect = DragDropEffects.Copy;
                    this.BackColor = ColorTranslator.FromHtml("#BFDBFE");
                    this.Invalidate();
                }
                else
                {
                    e.Effect = DragDropEffects.None;
                }
            }
        }

        private void ModernFilePanel_DragLeave(object sender, EventArgs e)
        {
            if (!fileLoaded && this.Enabled)
            {
                this.BackColor = ColorTranslator.FromHtml("#F8FAFC");
                this.Invalidate();
            }
        }

        private void ModernFilePanel_DragDrop(object sender, DragEventArgs e)
        {
            if (!this.Enabled) return;
            
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length == 1 && files[0].EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    onFileDropped?.Invoke(files[0]);
                }
            }
            
            if (!fileLoaded && this.Enabled)
            {
                this.BackColor = ColorTranslator.FromHtml("#F8FAFC");
                this.Invalidate();
            }
        }

        public void SetDropHandler(Action<string> handler)
        {
            this.onFileDropped = handler;
        }

        public void UpdateText(string text)
        {
            label.Text = text;
        }

        public void SetFileLoaded(string filename, int recordCount)
        {
            fileLoaded = true;
            this.BackColor = ColorTranslator.FromHtml("#DCFCE7");
            
            string displayName = filename.Length > 50 ? filename.Substring(0, 47) + "..." : filename;
            label.Text = $"{displayName}\n{recordCount} records loaded";
            label.ForeColor = ColorTranslator.FromHtml("#16A34A");
            label.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            
            this.Invalidate();
        }

        protected override void OnEnabledChanged(EventArgs e)
        {
            base.OnEnabledChanged(e);
            if (!Enabled)
            {
                this.BackColor = ColorTranslator.FromHtml("#F1F5F9");
                label.ForeColor = ColorTranslator.FromHtml("#94A3B8");
                this.Cursor = Cursors.Default;
                label.Cursor = Cursors.Default;
            }
            else
            {
                this.BackColor = ColorTranslator.FromHtml("#F8FAFC");
                label.ForeColor = ColorTranslator.FromHtml("#64748B");
                this.Cursor = Cursors.Hand;
                label.Cursor = Cursors.Hand;
            }
        }
    }

    public class ModernButton : Button
    {
        private bool isSecondary = false;
        private bool isRounded = false;

        public ModernButton()
        {
            this.FlatStyle = FlatStyle.Flat;
            this.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.Cursor = Cursors.Hand;
            UpdateStyle();
        }

        public void SetSecondary()
        {
            isSecondary = true;
            UpdateStyle();
        }

        public void SetRounded()
        {
            isRounded = true;
            this.Paint += ModernButton_Paint;
        }

        private void ModernButton_Paint(object sender, PaintEventArgs e)
        {
            if (isRounded)
            {
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                
                var rect = new Rectangle(0, 0, this.Width - 1, this.Height - 1);
                var path = GetRoundedRectanglePath(rect, 10);
                
                this.Region = new Region(path);
                
                using (var brush = new SolidBrush(this.BackColor))
                {
                    e.Graphics.FillPath(brush, path);
                }
                
                using (var pen = new Pen(this.BackColor, 2))
                {
                    e.Graphics.DrawPath(pen, path);
                }
                
                TextRenderer.DrawText(e.Graphics, this.Text, this.Font, rect, this.ForeColor, 
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            }
        }

        private System.Drawing.Drawing2D.GraphicsPath GetRoundedRectanglePath(Rectangle rect, int radius)
        {
            var path = new System.Drawing.Drawing2D.GraphicsPath();
            path.AddArc(rect.X, rect.Y, radius, radius, 180, 90);
            path.AddArc(rect.Right - radius, rect.Y, radius, radius, 270, 90);
            path.AddArc(rect.Right - radius, rect.Bottom - radius, radius, radius, 0, 90);
            path.AddArc(rect.X, rect.Bottom - radius, radius, radius, 90, 90);
            path.CloseFigure();
            return path;
        }

        private void UpdateStyle()
        {
            if (isSecondary)
            {
                this.BackColor = Color.White;
                this.ForeColor = ColorTranslator.FromHtml("#3B82F6");
                this.FlatAppearance.BorderColor = ColorTranslator.FromHtml("#3B82F6");
                this.FlatAppearance.BorderSize = 2;
                this.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml("#EFF6FF");
            }
            else
            {
                this.BackColor = ColorTranslator.FromHtml("#3B82F6");
                this.ForeColor = Color.White;
                this.FlatAppearance.BorderSize = 0;
                this.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml("#2563EB");
            }
        }

        protected override void OnEnabledChanged(EventArgs e)
        {
            base.OnEnabledChanged(e);
            if (!Enabled && !isSecondary)
            {
                this.BackColor = ColorTranslator.FromHtml("#CBD5E1");
                this.ForeColor = ColorTranslator.FromHtml("#94A3B8");
            }
            else
            {
                UpdateStyle();
            }
        }
    }
}