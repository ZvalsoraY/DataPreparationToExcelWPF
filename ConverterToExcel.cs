using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.Util;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel.Charts;

namespace DataPreparationToExcelWPF
{
    /// <summary>
    /// This class provides conversion of the allocated string array to Excel.
    /// </summary>
    public static class ConverterToExcel
    {
        public static List<string> list = new List<string>();
        /// <summary>
        /// This method invokes sorting and array transformation
        /// for each file in the selected list.
        /// </summary>
        public static void createExcelForListFiles(List<string> list, string mUnit)
        {
            foreach (var inputFileName in list)
            {
                createExcelFile(inputFileName, mUnit);
                createExcelFileYaxis(inputFileName, mUnit);
            }
        }
        /// <summary>
        /// This method groups the values along the z-axis
        /// and passes the data in groups for conversion to Excel.
        /// </summary>
        public static void createExcelFile(string fileName, string mUnit)
        {
            var listData = FileDoubleArrayList(fileName, mUnit);
            var query = listData.GroupBy(
            u => Math.Floor(u.ElementAt(2)),
            u => u);
            foreach (var result in query)
            {
                //generateExcel(result, fileName, $"B2:B", result.ElementAt(0).ElementAt(2));
                generateExcel(result, fileName.Substring(0, fileName.LastIndexOf('.')) + "_Z_axis", $"B2:B", result.ElementAt(0).ElementAt(2));
            }
        }
        /// <summary>
        /// This method groups the values along the y-axis
        /// and passes the data in groups for conversion to Excel.
        /// </summary>
        public static void createExcelFileYaxis(string fileName, string mUnit)
        {
            var listData = FileDoubleArrayList(fileName, mUnit, 2);
            var query = listData.GroupBy(
                u => Math.Floor(u.ElementAt(1)),
                u => u);
            foreach (var result in query)
            {
                generateExcel(result, fileName.Substring(0, fileName.LastIndexOf('.')) + "_Y_axis", $"C2:C", result.ElementAt(0).ElementAt(1));
            }
        }
        /// <summary>
        /// This method create IEnumerable IEnumerable double
        /// from file.
        /// </summary>
        public static IEnumerable<IEnumerable<double>> FileDoubleArrayList(string filePath, string mUnit, int parCol = 1)
        {
            var lines = File.ReadLines(filePath);
            if (mUnit == "m")
            {
                return lines.Select(line =>
              line.Split(new[] { ' ', '!' }, StringSplitOptions.RemoveEmptyEntries)
              .Select((s, index) => {
                  if (index < 3)
                      return Math.Round(double.Parse(s, NumberStyles.Any, CultureInfo.InvariantCulture) * 1000, 2, MidpointRounding.AwayFromZero);
                  return Math.Round(double.Parse(s, NumberStyles.Any, CultureInfo.InvariantCulture), 2, MidpointRounding.AwayFromZero);
              }))
                  .OrderBy(u => u.ElementAt(parCol));
            }
            else
            {
                //var lines = File.ReadLines(filePath);
                //return lines.Select(line =>
                //  line.Split(new[] { ' ', '!' }, StringSplitOptions.RemoveEmptyEntries).Select(s =>
                //      double.Parse(s, NumberStyles.Any, CultureInfo.InvariantCulture))).OrderBy(u => u.ElementAt(1));
                return lines.Select(line =>
                  line.Split(new[] { ' ', '!' }, StringSplitOptions.RemoveEmptyEntries).Select(s =>
                  Math.Round(double.Parse(s, NumberStyles.Any, CultureInfo.InvariantCulture), 2, MidpointRounding.AwayFromZero)
                      )).OrderBy(u => u.ElementAt(parCol));
            }
        }
        /// <summary>
        /// This method create Excel file with chart.
        /// </summary>
        public static void generateExcel(IEnumerable<IEnumerable<double>> inputArray, string fileNames, string columnSel, double axisCoordinateData = 0.0)
        {
            //string parth = $"D:\\test\\" + fileNames + $" dist {axisCoordinateData.ToString()}.xlsx";
            //string parth = fileNames.Substring(0, fileNames.LastIndexOf('.')) + $" dist {axisCoordinateData.ToString()}.xlsx";
            string parth = fileNames + $" dist {axisCoordinateData.ToString()}.xlsx";
            using (var stream = new FileStream(parth, FileMode.Create, FileAccess.ReadWrite))
            {
                var wb = new XSSFWorkbook();
                //var wb = new ;
                var sheet = wb.CreateSheet($"Hsum dist {axisCoordinateData.ToString()}");
                //creating cell style for header
                var bStylehead = wb.CreateCellStyle();
                bStylehead.BorderBottom = BorderStyle.Thin;
                bStylehead.BorderLeft = BorderStyle.Thin;
                bStylehead.BorderRight = BorderStyle.Thin;
                bStylehead.BorderTop = BorderStyle.Thin;
                bStylehead.Alignment = HorizontalAlignment.Center;
                bStylehead.VerticalAlignment = VerticalAlignment.Center;
                bStylehead.FillBackgroundColor = HSSFColor.Green.Index;
                //var cellStyle =
                //var cellStyle = CreateCellStyleForHeader(wb);
                var Drawing = sheet.CreateDrawingPatriarch();
                //IClientAnchor anchor = Drawing.CreateAnchor(0, 0, 0, 0, 8, 1, 18, 16);
                var anchor = Drawing.CreateAnchor(0, 0, 0, 0, 8, 1, 23, 16);
                var chart = Drawing.CreateChart(anchor);
                //IChart chart = Drawing.CreateChart(anchor);
                IChartAxis bottomAxis = chart.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
                IChartAxis leftAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);

                var chartData =
                        chart.ChartDataFactory.CreateLineChartData<double, double>();
                //var lenCellRange = inputArray.Count + 1;
                var lenCellRange = inputArray.Count() + 1;
                //var columnLenCellRange = columnSel + $"{lenCellRange}";
                //IChartDataSource<double> xs = DataSources.FromNumericCellRange(sheet, CellRangeAddress.ValueOf($"B2:B{lenCellRange}"));
                IChartDataSource<double> xs = DataSources.FromNumericCellRange(sheet, CellRangeAddress.ValueOf(columnSel + $"{lenCellRange}"));
                IChartDataSource<double> ys = DataSources.FromNumericCellRange(sheet, CellRangeAddress.ValueOf($"G2:G{lenCellRange}"));
                //IChartDataSource<double> ys = DataSources.FromNumericCellRange(sheet, CellRangeAddress.ValueOf("G2:G20"));

                var series = chartData.AddSeries(xs, ys);
                series.SetTitle("Hsum");
                //chart.GetOrCreateLegend();
                chart.Plot(chartData, bottomAxis, leftAxis);

                var row = sheet.CreateRow(0);

                row.CreateCell(0, CellType.String).SetCellValue("x");
                row.CreateCell(1, CellType.String).SetCellValue("y");
                row.CreateCell(2, CellType.String).SetCellValue("z");
                row.CreateCell(3, CellType.String).SetCellValue("Hx");
                row.CreateCell(4, CellType.String).SetCellValue("Hy");
                row.CreateCell(5, CellType.String).SetCellValue("Hz");
                row.CreateCell(6, CellType.String).SetCellValue("Hsum");
                //row.Cells[0].CellStyle = bStylehead;
                //row.RowStyle = bStylehead;

                //filling the data
                var rowsCounter = 1;
                foreach (var rowData in inputArray)
                {
                    var rowD = sheet.CreateRow(rowsCounter++);
                    var dCounter = 0;
                    foreach (var d in rowData)
                    {
                        //rowD.CreateCell(dCounter++, CellType.Numeric).SetCellValue(Double.Parse(d.ToString().Replace(@".", @",")));
                        //CultureInfo.InvariantCulture
                        rowD.CreateCell(dCounter++, CellType.Numeric).SetCellValue(d);
                    }
                }
                wb.Write(stream);
                wb.Close();
            }
        }
    }
}
