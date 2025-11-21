// Models/InTrayRecord.cs

using CsvHelper.Configuration.Attributes;

namespace ADMerger.Models
{
    public class InTrayRecord
    {
        [Name("Received on")]
        public string ReceivedOn { get; set; }
       
        [Name("Student No.")]
        public string StudentNo { get; set; }
       
        [Name("Name")]
        public string Name { get; set; }
    }
}