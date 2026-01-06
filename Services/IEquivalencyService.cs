// Services/IEquivalencyService.cs
// © Ricki Angel 2026 | TechAngelX

using System.Collections.Generic;
using ADMerger.Models;

namespace ADMerger.Services
{
    public interface IEquivalencyService
    {
        void LoadEquivalencies();
        DegreeEquivalency GetEquivalency(string country);
        Dictionary<string, DegreeEquivalency> GetAllEquivalencies();
        int Count { get; }
    }
}
