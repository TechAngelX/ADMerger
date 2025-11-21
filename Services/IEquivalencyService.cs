// Services/IEquivalencyService.cs

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