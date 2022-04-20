using ClosedXML.Excel;
using ExporterAPI.Model;
using System.Data;

namespace ExporterAPI.Helpers
{
    public static class ExporterExtensions
    {
        public static IXLPivotTable AddPivotTable(this IXLWorksheet sheet, int StartRow, IXLRange sc, IEnumerable<string> rows, IEnumerable<string> columns, IEnumerable<string> measures)
        {
            string pivotTableName = "PivotTable" + sheet.PivotTables.Count() + 1;
            var pt = sheet.PivotTables.Add(pivotTableName, sheet.Cell(StartRow, 1), sc);
            rows.ToList().ForEach(r => pt.RowLabels.Add(r));
            columns.ToList().ForEach(c => pt.ColumnLabels.Add(c));
            measures.ToList().ForEach(m => pt.Values.Add(m));
            return pt;
        }
        public static IXLPivotTable AddPivotTable(this IXLWorksheet sheet, int StartRow, IXLRange sc, List<Definition> defs)
        {
            string pivotTableName = "PivotTable" + sheet.PivotTables.Count() + 1;
            var pt = sheet.PivotTables.Add(pivotTableName, sheet.Cell(StartRow, 1), sc);
            pt.AutofitColumns = true;
            foreach (Definition def in defs)
            {
                switch (def.Role)
                {
                    case "Row":
                        pt.RowLabels.Add(def.SourceName);
                        pt.RowLabels.Get(def.SourceName).CustomName = def.DisplayName;
                        
                        break;
                    case "Column":
                        pt.ColumnLabels.Add(def.SourceName);
                        pt.ColumnLabels.Get(def.SourceName).CustomName = def.DisplayName;
                        
                        break;
                    case "Measure":
                        pt.Values.Add(def.SourceName);
                        if (!String.IsNullOrEmpty(def.Format))
                            pt.Values.Get(def.SourceName).NumberFormat.Format = def.Format;
                        pt.Values.Get(def.SourceName).CustomName = def.DisplayName;
                        
                        break;
                }
            }
            pt.SortFieldsAtoZ = true;
            pt.AutofitColumns = true;
            pt.RefreshDataOnOpen = true;
            return pt;
        }
        public static IXLRange AddDataSourceSheet(this XLWorkbook wb, DataTable dt)
        {
            var sc = wb.AddWorksheet(dt);
            sc.Hide();
            return sc.Table(dt.TableName).AsRange();
        }
        public static IXLRange AddDataSourceSheet(this XLWorkbook wb, ExportData ed)
        {
            var sc = wb.AddWorksheet("Datasource");

            IXLTable tab = sc.Cell(1, 1).InsertTable(ed.Data,"ExportData",true   );
            foreach (Definition def in ed.Definitions)
            {
                XLDataType fieldType = GetXLDatatypeFromDef(def.DataType);
                string format = CleanFormatString(def.Format);
                tab.Field(def.SourceName).DataCells.DataType = fieldType;
                if (!String.IsNullOrEmpty(def.Format))
                    tab.Field(def.SourceName).DataCells.Style.NumberFormat.Format = def.Format;
                

            }
            sc.Hide();
            return sc.Table(ed.Data.TableName).AsRange();
        }
        public static IXLCell InsertTitle(this IXLWorksheet sheet, string title, int rowIndex)
        {
            var c = sheet.Cell(rowIndex, 1);
            c.Value = title;
            c.Style.Font.SetFontSize(26);
            c.Style.Font.SetBold(true);
            sheet.Row(rowIndex).Merge();
            return c;
        }
        public static string CleanFormatString(string format)
        {
            return format.Split(';')[0].Replace("\"", "");
        }
        public static IXLCell InsertSubTitle(this IXLWorksheet sheet, string subtitle, int rowIndex)
        {
            var c = sheet.Cell(rowIndex, 1);
            c.Value = subtitle;
            c.Style.Font.SetFontSize(20);
            sheet.Row(rowIndex).Merge();
            return c;
        }
        public static IXLTable InsertTable(this IXLWorksheet sheet, DataTable dt, int rowIndex)
        {
            var tab = sheet.Cell(rowIndex, 1).InsertTable(dt);
            tab.Theme = XLTableTheme.TableStyleLight2;
            tab.SetShowAutoFilter(false);
            tab.SetShowTotalsRow(false);
            return tab;
        }
        public static IXLTable InsertTable(this IXLWorksheet sheet, ExportData ed, int rowIndex)
        {

            var tab = sheet.Cell(rowIndex, 1).InsertTable(ed.Data);
            tab.Theme = XLTableTheme.TableStyleLight2;
            tab.SetShowAutoFilter(true);
            tab.SetShowTotalsRow(true);
            foreach (Definition def in ed.Definitions)
            {
                XLDataType fieldType = GetXLDatatypeFromDef(def.DataType);
                tab.Field(def.SourceName).DataCells.DataType = fieldType;
                if (!String.IsNullOrEmpty(def.Format))
                    tab.Field(def.SourceName).DataCells.Style.NumberFormat.Format = def.Format;

                if (fieldType == XLDataType.Number)
                    tab.Field(def.SourceName).TotalsRowFunction = XLTotalsRowFunction.Sum;

                tab.Field(def.SourceName).Name = def.DisplayName;
            }
            tab.Field(0).TotalsRowLabel = "Total";
            
            return tab;
        }
        private static XLDataType GetXLDatatypeFromDef(string def)
        {
            XLDataType res = XLDataType.Text;
            switch (def)
            {
                case "bool":
                    res = XLDataType.Boolean;
                    break;
                case "dateTime":
                    res = XLDataType.DateTime;
                    break;
                case "duration":
                    res = XLDataType.TimeSpan;
                    break;
                case "integer":
                    res = XLDataType.Number;
                    break;
                case "numeric":
                    res = XLDataType.Number;
                    break;

                default:
                    res = XLDataType.Text;
                    break;
            }

            return res;
        }
    }
}
