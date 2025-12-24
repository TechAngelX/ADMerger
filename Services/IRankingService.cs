// Services/IRankingService.cs

using System.Collections.Generic;
using System.Threading.Tasks;

namespace ADMerger.Services
{
    public interface IRankingService
    {
        Task LoadRankingsAsync();
        string GetRanking(string institutionName);
        int Count { get; }
        IReadOnlyList<string> GetAllInstitutionNames();
    }
}
