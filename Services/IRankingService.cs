// Services/IRankingService.cs

using System.Collections.Generic;

namespace ADMerger.Services
{
    public interface IRankingService
    {
        void LoadRankings();
        string GetRanking(string institutionName);
        int Count { get; }
        IReadOnlyList<string> GetAllInstitutionNames();
    }
}