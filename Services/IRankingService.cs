// Services/IRankingService.cs
// © Ricki Angel 2026 | TechAngelX

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
