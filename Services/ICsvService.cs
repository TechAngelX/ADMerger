// Services/ICsvService.cs
// © Ricki Angel 2026 | TechAngelX

using System.Collections.Generic;
using ADMerger.Models;

namespace ADMerger.Services
{
    public interface ICsvService
    {
        List<InTrayRecord> LoadInTrayRecords(string filePath);
        List<ApplicationRecord> LoadApplicationRecords(string filePath);
        List<string> GenerateOutputFiles(List<OutputRecord> data, string outputFolderPath, List<string>? additionalFields = null);
    }
}
