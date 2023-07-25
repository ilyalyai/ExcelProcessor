using EPPlusSamples;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;

namespace EpPlus
{
    public class EpPlusClass
    {
        private readonly string fileName;

        public EpPlusClass(string fileName)
        {
            this.fileName = fileName;
        }

        public string Run()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo templateFile = FileUtil.GetFileInfo(fileName);
            FileInfo newFile = FileUtil.GetCleanFileInfo(templateFile.FullName.Replace(templateFile.Extension, "") + "new" + templateFile.Extension);

            using (ExcelPackage package = new ExcelPackage(newFile, templateFile))
            {
                //Open the first worksheet
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                worksheet.InsertRow(5, 2);

                worksheet.Cells["A5"].Value = "12010";
                worksheet.Cells["B5"].Value = "Drill";
                worksheet.Cells["C5"].Value = 20;
                worksheet.Cells["D5"].Value = 8;

                worksheet.Cells["A6"].Value = "12011";
                worksheet.Cells["B6"].Value = "Crowbar";
                worksheet.Cells["C6"].Value = 7;
                worksheet.Cells["D6"].Value = 23.48;

                worksheet.Cells["E2:E6"].FormulaR1C1 = "RC[-2]*RC[-1]";

                var name = worksheet.Names.Add("SubTotalName", worksheet.Cells["C7:E7"]);
                name.Style.Font.Italic = true;
                name.Formula = "SUBTOTAL(9,C2:C6)";

                //Format the new rows
                worksheet.Cells["C5:C6"].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["D5:E6"].Style.Numberformat.Format = "#,##0.00";

                var chart = worksheet.Drawings.AddPieChart("PieChart", ePieChartType.Pie3D);

                chart.Title.Text = "Total";
                //From row 1 colum 5 with five pixels offset
                chart.SetPosition(0, 0, 5, 5);
                chart.SetSize(600, 300);

                ExcelAddress valueAddress = new ExcelAddress(2, 5, 6, 5);
                var ser = chart.Series.Add(valueAddress.Address, "B2:B6") as ExcelPieChartSerie;
                chart.DataLabel.ShowCategory = true;
                chart.DataLabel.ShowPercent = true;

                chart.Legend.Border.LineStyle = eLineStyle.Solid;
                chart.Legend.Border.Fill.Style = eFillStyle.SolidFill;
                chart.Legend.Border.Fill.Color = Color.DarkBlue;

                //Set the chart style to match the preset style for 3D pie charts.
                chart.StyleManager.SetChartStyle(ePresetChartStyle.Pie3dChartStyle3);

                //Switch the PageLayoutView back to normal
                worksheet.View.PageLayoutView = false;
                // save our new workbook and we are done!
                package.Save();
            }

            return newFile.FullName;
        }
    }
}