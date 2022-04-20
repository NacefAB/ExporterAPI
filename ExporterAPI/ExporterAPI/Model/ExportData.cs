using System.Data;

namespace ExporterAPI.Model
{
    public class ExportData
    {
        public string Title { get; set; }
        public DataTable Data { get; set; }
        public List<Definition> Definitions { get; set; }
        public string Filters { get; set; }
    }
}
