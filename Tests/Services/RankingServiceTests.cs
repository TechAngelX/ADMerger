// Tests/Services/RankingServiceTests.cs

using System.Collections.Generic;
using System.Threading.Tasks;
using Xunit;
using Moq;
using ADMerger.Services;

namespace ADMerger.Tests.Services
{
    public class RankingServiceTests
    {
        private readonly RankingService _service;
        private readonly Mock<IInstitutionMatchingService> _mockMatchingService;

        public RankingServiceTests()
        {
            _mockMatchingService = new Mock<IInstitutionMatchingService>();
            _service = new RankingService(_mockMatchingService.Object);
        }

        #region Initialization Tests

        [Fact]
        public async Task LoadRankingsAsync_LoadsInstitutionNames()
        {
            await _service.LoadRankingsAsync();
            
            var names = _service.GetAllInstitutionNames();
            Assert.NotEmpty(names);
            Assert.True(names.Count > 0);
        }

        [Fact]
        public void Count_ReturnsNumberOfRankings()
        {
            var count = _service.Count;
            Assert.True(count > 0);
        }

        #endregion

        #region Null/Empty Input Tests

        [Fact]
        public void GetRanking_NullInput_ReturnsNR()
        {
            var result = _service.GetRanking(null);
            Assert.Equal("NR", result);
        }

        [Fact]
        public void GetRanking_EmptyString_ReturnsNR()
        {
            var result = _service.GetRanking("");
            Assert.Equal("NR", result);
        }

        [Fact]
        public void GetRanking_WhitespaceOnly_ReturnsNR()
        {
            var result = _service.GetRanking("   ");
            Assert.Equal("NR", result);
        }

        #endregion

        #region Exact Match Tests

        [Fact]
        public void GetRanking_ExactMatch_ReturnsRank()
        {
            // Use a university you know is in RankingData
            var result = _service.GetRanking("University of Oxford");
            Assert.NotEqual("NR", result);
            Assert.NotNull(result);
        }

        [Fact]
        public void GetRanking_ExactMatchCaseInsensitive_ReturnsRank()
        {
            // Assuming your dictionary uses case-insensitive keys
            var result = _service.GetRanking("UNIVERSITY OF OXFORD");
            Assert.NotEqual("NR", result);
        }

        #endregion

        #region Mapping Tests

        [Fact]
        public void GetRanking_AbbreviationUCL_ReturnsRank()
        {
            // Assuming MappingData has: "UCL" -> "University College London"
            var result = _service.GetRanking("UCL");
            Assert.NotEqual("NR", result);
        }

        [Fact]
        public void GetRanking_AbbreviationMIT_ReturnsRank()
        {
            // Assuming MappingData has: "MIT" -> "Massachusetts Institute of Technology"
            var result = _service.GetRanking("MIT");
            Assert.NotEqual("NR", result);
        }

        #endregion

        #region Fuzzy Match Tests

        [Fact]
        public async Task GetRanking_FuzzyMatch_UsesMatchingService()
        {
            await _service.LoadRankingsAsync();
            
            // Setup mock to return a known university
            _mockMatchingService
                .Setup(x => x.FindBestMatch(It.IsAny<string>(), It.IsAny<List<string>>()))
                .Returns("University of Oxford");

            var result = _service.GetRanking("Oxfrod University"); // Typo
            
            // Should have called the matching service
            _mockMatchingService.Verify(
                x => x.FindBestMatch(It.IsAny<string>(), It.IsAny<List<string>>()), 
                Times.Once
            );
            
            Assert.NotEqual("NR", result);
        }

        [Fact]
        public async Task GetRanking_NoFuzzyMatch_ReturnsNR()
        {
            await _service.LoadRankingsAsync();
            
            // Setup mock to return null (no match)
            _mockMatchingService
                .Setup(x => x.FindBestMatch(It.IsAny<string>(), It.IsAny<List<string>>()))
                .Returns((string?)null);

            var result = _service.GetRanking("University of Mars");
            
            Assert.Equal("NR", result);
        }

        #endregion

        #region Integration-Style Tests

        [Fact]
        public void GetRanking_TopUniversities_ReturnValidRanks()
        {
            // Test a few known top universities
            var oxford = _service.GetRanking("University of Oxford");
            var cambridge = _service.GetRanking("University of Cambridge");
            var imperial = _service.GetRanking("Imperial College London");

            Assert.NotEqual("NR", oxford);
            Assert.NotEqual("NR", cambridge);
            Assert.NotEqual("NR", imperial);
        }

        [Fact]
        public async Task GetRanking_WithLeadingTrailingSpaces_TrimsAndMatches()
        {
            await _service.LoadRankingsAsync();
            
            var result = _service.GetRanking("  University of Oxford  ");
            Assert.NotEqual("NR", result);
        }

        #endregion

        #region Edge Cases

        [Fact]
        public void GetRanking_UnknownUniversity_ReturnsNR()
        {
            _mockMatchingService
                .Setup(x => x.FindBestMatch(It.IsAny<string>(), It.IsAny<List<string>>()))
                .Returns((string?)null);

            var result = _service.GetRanking("Fake University That Does Not Exist");
            Assert.Equal("NR", result);
        }

        [Fact]
        public async Task GetAllInstitutionNames_ReturnsReadOnlyList()
        {
            await _service.LoadRankingsAsync();
            
            var names = _service.GetAllInstitutionNames();
            
            Assert.NotNull(names);
            Assert.IsAssignableFrom<IReadOnlyList<string>>(names);
        }

        #endregion
    }
}
