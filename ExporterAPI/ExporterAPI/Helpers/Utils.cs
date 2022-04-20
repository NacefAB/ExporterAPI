using ClosedXML.Excel;
using ExporterAPI.Model;
using System.Data;
using System.Text;

namespace ExporterAPI.Helpers
{
    public class Utils
    {
        public static string ExportDataToXlsx(ExportData ed)
        {

            XLWorkbook wb=new XLWorkbook();
            int currentRowIndex = 1;
            ed.Data.TableName = "ExportData";
            var sheet = wb.AddWorksheet("Data");
            if (!String.IsNullOrEmpty(ed.Title))
            {
                sheet.InsertTitle(ed.Title, currentRowIndex);
                currentRowIndex += 2;
            }
            if (!String.IsNullOrEmpty(ed.Filters))
            {
                sheet.InsertSubTitle("Filters",currentRowIndex);
                currentRowIndex++;
                sheet.InsertTable(generateFiltersSummary(ed.Filters), currentRowIndex);
                currentRowIndex +=4;
            }
            if (isPivotExport(ed.Definitions))
            {
                IXLRange src=wb.AddDataSourceSheet(ed);
                sheet.AddPivotTable(currentRowIndex, src, ed.Definitions);

            }
            else
            {
                sheet.InsertTable(ed, currentRowIndex);
            }
            string filename = GetTempFile();
            sheet.Columns().AdjustToContents();
            
            wb.SaveAs(filename, false);

            return filename;
        }
        public static bool isPivotExport(List<Definition> defs)
        {
            if(defs.Any(d=>d.Role== "Column"))
                return true;
            return false;
        }
        public static DataTable generateFiltersSummary(string filters)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Name",typeof(string));
            dt.Columns.Add("Value",typeof(string));
            foreach (var fil in filters.Split(';'))
                dt.Rows.Add(fil.Split(':'));
            dt.TableName = "Filters";
            return dt;
        }
        public static string GetTempFile()
        {
            string exportDir=GetExportFolder();
            string tempFile = Path.Combine(exportDir, Guid.NewGuid() + ".xlsx");
            Directory.CreateDirectory(exportDir);
            //Delete generated files older than 30 min 
            exportDirCleanup(exportDir);
            return tempFile;
        }
        public static string GetExportFolder()
        {
            return Path.Combine(Path.GetTempPath(), "ExporterAPI\\");
        }
        public static string CleanSpecialChars(String text)
        {
            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < text.Length; i++)
            {
                if (char.IsLetterOrDigit(text[i]))
                {
                    sb.Append(text[i]);
                }
            }
            return sb.ToString();
        }
        private static void exportDirCleanup(string exportDir)
        {
            string[] files = Directory.GetFiles(exportDir);

            foreach (string file in files)
            {
                FileInfo fi = new FileInfo(file);
                if (fi.LastWriteTime < DateTime.Now.AddMinutes(-30))
                    fi.Delete();
            }
        }
    }
}
