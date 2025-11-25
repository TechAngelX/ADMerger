// Services/ICsvService.cs

using System.Collections.Generic;
using ADMerger.Models;

namespace ADMerger.Services
{
    public interface ICsvService
    {
        List<InTrayRecord> LoadInTrayRecords(string filePath);
        List<ApplicationRecord> LoadApplicationRecords(string filePath);
        string GenerateOutputFiles(List<OutputRecord> data, string outputFolderPath);
    }
}