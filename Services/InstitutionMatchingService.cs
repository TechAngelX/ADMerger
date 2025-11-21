// Services/InstitutionMatchingService.cs

using System;
using System.Collections.Generic;
using ADMerger.Utilities;

namespace ADMerger.Services
{
    public class InstitutionMatchingService : IInstitutionMatchingService
    {
        private const int MinimumMatchThreshold = 60;
        
        public string FindBestMatch(string searchName, List<string> candidateNames)
        {
            if (string.IsNullOrWhiteSpace(searchName))
                return null;
            
            string normalizedSearch = TextNormalizer.NormalizeInstitutionName(searchName);
            var searchTerms = TextNormalizer.ExtractKeyTerms(normalizedSearch);
            
            string bestMatch = null;
            int bestScore = 0;
            
            foreach (var candidateName in candidateNames)
            {
                string normalizedCandidate = TextNormalizer.NormalizeInstitutionName(candidateName);
                int score = CalculateMatchScore(normalizedSearch, normalizedCandidate, searchTerms);
                
                if (score > bestScore && score >= MinimumMatchThreshold)
                {
                    bestScore = score;
                    bestMatch = candidateName;
                }
            }
            
            return bestMatch;
        }
        
        public int CalculateMatchScore(string search, string candidate, List<string> searchTerms)
        {
            int score = 0;
            
            if (search == candidate)
                return 100;
            
            if (candidate.Contains(search))
                score += 80;
            
            if (search.Contains(candidate))
                score += 70;
            
            int matchedTerms = 0;
            foreach (var term in searchTerms)
            {
                if (candidate.Contains(term))
                    matchedTerms++;
            }
            
            if (searchTerms.Count > 0)
            {
                int termScore = (matchedTerms * 100) / searchTerms.Count;
                score = Math.Max(score, termScore);
            }
            
            score = Math.Max(score, CheckSpecialCases(search, candidate));
            
            return score;
        }
        
        private int CheckSpecialCases(string search, string candidate)
        {
            if ((search.Contains("ucl") || search.Contains("university college london")) && 
                candidate.Contains("university college london"))
                return 95;
            
            if ((search.Contains("oxford") && candidate.Contains("oxford")) ||
                (search.Contains("cambridge") && candidate.Contains("cambridge")) ||
                (search.Contains("mit") && candidate.Contains("massachusetts institute")) ||
                (search.Contains("caltech") && candidate.Contains("california institute")))
                return 90;
            
            return 0;
        }
    }
}