// Services/IInstitutionMatchingService.cs
// © Ricki Angel 2026 | TechAngelX

using System.Collections.Generic;

namespace ADMerger.Services
{
    public interface IInstitutionMatchingService
    {
        string FindBestMatch(string searchName, List<string> candidateNames);
        int CalculateMatchScore(string searchName, string candidateName, List<string> searchTerms);
    }
}
