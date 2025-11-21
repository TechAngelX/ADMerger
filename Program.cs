// Program.cs

using System;
using System.Windows.Forms;
using ADMerger.Services;

namespace ADMerger
{
    internal class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            var csvService = new CsvService();
            var equivalencyService = new EquivalencyService();
            var matchingService = new InstitutionMatchingService();
            var rankingService = new RankingService(matchingService);
            var gradeService = new GradeClassificationService(equivalencyService);
            
            var mainForm = new MainForm(
                csvService,
                equivalencyService,
                rankingService,
                gradeService
            );
            
            Application.Run(mainForm);
        }
    }
}