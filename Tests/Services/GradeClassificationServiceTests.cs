// Tests/Services/GradeClassificationServiceTests.cs
// © Ricki Angel 2026 | TechAngelX


using Xunit;
using Moq;
using ADMerger.Services;
using ADMerger.Models;

namespace ADMerger.Tests.Services
{
    public class GradeClassificationServiceTests
    {
        private readonly GradeClassificationService _service;
        private readonly Mock<IEquivalencyService> _mockEquivService;

        public GradeClassificationServiceTests()
        {
            _mockEquivService = new Mock<IEquivalencyService>();
            _service = new GradeClassificationService(_mockEquivService.Object);
        }

        #region Text Classification Tests

        [Fact]
        public void TextKeywords_FirstClass_ReturnsFirst()
        {
            var result = _service.DetermineUKClassification(null, "Overall classification was First Class Honours", "United Kingdom", "BSc");
            Assert.Equal("1.0", result);
        }

        [Fact]
        public void TextKeywords_UpperSecond_Returns21()
        {
            var result = _service.DetermineUKClassification(null, "Degree awarded: Upper Second Class (2:1)", "United Kingdom", "BSc");
            Assert.Equal("2.1", result);
        }

        [Fact]
        public void TextKeywords_LowerSecond_Returns22()
        {
            var result = _service.DetermineUKClassification(null, "Lower Second Class Honours", "United Kingdom", "BSc");
            Assert.Equal("2.2", result);
        }

        [Fact]
        public void TextKeywords_ThirdClass_ReturnsThird()
        {
            var result = _service.DetermineUKClassification(null, "Third Class Honours awarded", "United Kingdom", "BSc");
            Assert.Equal("3.0", result);
        }

        [Fact]
        public void TextKeywords_SummaCumLaude_ReturnsFirst()
        {
            var result = _service.DetermineUKClassification(null, "Graduated Summa Cum Laude", "United States", "Bachelor");
            Assert.Equal("1.0", result);
        }

        [Fact]
        public void TextKeywords_Distinction_ReturnsFirst()
        {
            var result = _service.DetermineUKClassification(null, "Pass with Distinction", "Australia", "Bachelor");
            Assert.Equal("1.0", result);
        }

        [Fact]
        public void TextKeywords_Merit_Returns21()
        {
            var result = _service.DetermineUKClassification(null, "Pass with Merit", "Australia", "Bachelor");
            Assert.Equal("2.1", result);
        }

        #endregion

        #region UK Grading Tests

        [Fact]
        public void UKGrade_75Percent_ReturnsFirst()
        {
            var result = _service.DetermineUKClassification(null, "Grade: 75%", "United Kingdom", "BSc");
            Assert.Equal("1.0", result);
        }

        [Fact]
        public void UKGrade_65Percent_Returns21()
        {
            var result = _service.DetermineUKClassification(null, "Average 65%", "United Kingdom", "BSc");
            Assert.Equal("2.1", result);
        }

        [Fact]
        public void UKGrade_55Percent_Returns22()
        {
            var result = _service.DetermineUKClassification(null, "Final grade 55", "United Kingdom", "BSc");
            Assert.Equal("2.2", result);
        }

        [Fact]
        public void UKGrade_45Percent_ReturnsThird()
        {
            var result = _service.DetermineUKClassification(null, "Overall 45%", "United Kingdom", "BSc");
            Assert.Equal("3.0", result);
        }

        [Fact]
        public void UKGrade_BelowPass_ReturnsFail()
        {
            var result = _service.DetermineUKClassification(null, "Grade 35", "United Kingdom", "BSc");
            Assert.Equal("Fail", result);
        }

        #endregion

        #region US GPA Tests

        [Fact]
        public void USGPA_3Point9_ReturnsFirst()
        {
            var usEquiv = new DegreeEquivalency 
            { 
                Country = "United States",
                First = "3.7",
                SecondUpper = "3.3",
                SecondLower = "3.0"
            };
            _mockEquivService.Setup(x => x.GetEquivalency("United States")).Returns(usEquiv);

            var result = _service.DetermineUKClassification(null, "GPA: 3.9/4.0", "United States", "Bachelor");
            Assert.Equal("1.0", result);
        }

        [Fact]
        public void USGPA_3Point5_Returns21()
        {
            var usEquiv = new DegreeEquivalency 
            { 
                Country = "United States",
                First = "3.7",
                SecondUpper = "3.3",
                SecondLower = "3.0"
            };
            _mockEquivService.Setup(x => x.GetEquivalency("United States")).Returns(usEquiv);

            var result = _service.DetermineUKClassification(null, "CGPA: 3.5", "United States", "Bachelor");
            Assert.Equal("2.1", result);
        }

        #endregion

        #region German/Austrian Reverse Scale Tests

        [Fact]
        public void GermanGrade_1Point3_ReturnsFirst()
        {
            var germanEquiv = new DegreeEquivalency 
            { 
                Country = "Germany",
                First = "1.5",
                SecondUpper = "2.0",
                SecondLower = "3.0"
            };
            _mockEquivService.Setup(x => x.GetEquivalency("Germany")).Returns(germanEquiv);

            var result = _service.DetermineUKClassification(null, "Final grade 1.3", "Germany", "Bachelor");
            Assert.Equal("1.0", result);
        }

        [Fact]
        public void GermanGrade_1Point8_Returns21()
        {
            var germanEquiv = new DegreeEquivalency 
            { 
                Country = "Germany",
                First = "1.5",
                SecondUpper = "2.0",
                SecondLower = "3.0"
            };
            _mockEquivService.Setup(x => x.GetEquivalency("Germany")).Returns(germanEquiv);

            var result = _service.DetermineUKClassification(null, "Grade: 1.8", "Germany", "Bachelor");
            Assert.Equal("2.1", result);
        }

        #endregion

        #region Special Cases - Lancaster/Glasgow

        [Fact]
        public void LancasterGrade_18_ReturnsFirst()
        {
            var result = _service.DetermineUKClassification(null, "Lancaster University grade: 18", "United Kingdom", "BSc");
            Assert.Equal("1.0", result);
        }

        [Fact]
        public void GlasgowGrade_15_Returns21()
        {
            var result = _service.DetermineUKClassification(null, "Glasgow University final: 15", "United Kingdom", "BSc");
            Assert.Equal("2.1", result);
        }

        #endregion

        #region Special Cases - Italy

        [Fact]
        public void ItalyGrade_110_ReturnsFirst()
        {
            var result = _service.DetermineUKClassification(null, "Grade: 110/110", "Italy", "Laurea");
            Assert.Equal("1.0", result);
        }

        [Fact]
        public void ItalyGrade_107_Returns21()
        {
            var result = _service.DetermineUKClassification(null, "Final mark: 107", "Italy", "Laurea");
            Assert.Equal("2.1", result);
        }

        [Fact]
        public void Italy30Scale_29_ReturnsFirst()
        {
            var result = _service.DetermineUKClassification(null, "Average: 29/30", "Italy", "Laurea");
            Assert.Equal("1.0", result);
        }

        #endregion

        #region Masters Shield

        [Fact]
        public void MastersQualification_NoUndergradKeywords_ReturnsMasters()
        {
            var result = _service.DetermineUKClassification(null, "Grade: 65", "United Kingdom", "Masters in Computer Science");
            Assert.Equal("Masters", result);
        }

        [Fact]
        public void MastersQualification_WithUndergradKeywords_ReturnsClassification()
        {
            var result = _service.DetermineUKClassification(null, "Previous degree: 2.1", "United Kingdom", "Masters in AI");
            Assert.Equal("2.1", result);
        }

        #endregion

        #region Edge Cases

        [Fact]
        public void EmptyNote_ReturnsUnknown()
        {
            var result = _service.DetermineUKClassification("75", "", "United Kingdom", "BSc");
            Assert.Equal("??", result);
        }

        [Fact]
        public void NullNote_ReturnsUnknown()
        {
            var result = _service.DetermineUKClassification("75", null, "United Kingdom", "BSc");
            Assert.Equal("??", result);
        }

        [Fact]
        public void NoteWithUnderscores_ParsesCorrectly()
        {
            var result = _service.DetermineUKClassification(null, "Grade__75%__First_Class", "United Kingdom", "BSc");
            Assert.Equal("1.0", result);
        }

        [Fact]
        public void RequirementsInNote_IgnoresRequirements()
        {
            // Should not classify based on "requires 2.1" but on actual grade
            var result = _service.DetermineUKClassification(null, "Requires 2.1 for entry. Student achieved First Class.", "United Kingdom", "BSc");
            Assert.Equal("1.0", result);
        }

        #endregion

        #region Priority Keyword Extraction

        [Fact]
        public void PriorityKeyword_Average_ExtractsCorrectly()
        {
            var result = _service.DetermineUKClassification(null, "Various grades but average: 72", "United Kingdom", "BSc");
            Assert.Equal("1.0", result);
        }

        [Fact]
        public void PriorityKeyword_CGPA_ExtractsCorrectly()
        {
            var usEquiv = new DegreeEquivalency 
            { 
                Country = "United States",
                First = "3.7",
                SecondUpper = "3.3",
                SecondLower = "3.0"
            };
            _mockEquivService.Setup(x => x.GetEquivalency("United States")).Returns(usEquiv);

            var result = _service.DetermineUKClassification(null, "CGPA: 3.8 out of 4.0", "United States", "Bachelor");
            Assert.Equal("1.0", result);
        }

        #endregion

        #region Interface Method Tests

        [Fact]
        public void ParseUKGradeText_FirstClass_ReturnsFirst()
        {
            var result = _service.ParseUKGradeText("First Class Honours");
            Assert.Equal("1.0", result);
        }

        [Fact]
        public void ParseGradeValue_Percentage_ReturnsNumeric()
        {
            var result = _service.ParseGradeValue("75%");
            Assert.Equal(75.0, result);
        }

        [Fact]
        public void ParseGradeValue_Fraction_ReturnsNumerator()
        {
            var result = _service.ParseGradeValue("3.8/4.0");
            Assert.Equal(3.8, result);
        }

        [Fact]
        public void ParseGradeValue_InvalidInput_ReturnsNull()
        {
            var result = _service.ParseGradeValue("Not A Number");
            Assert.Null(result);
        }

        #endregion
    }
}

