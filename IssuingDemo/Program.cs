using CsvHelper;
using CsvHelper.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace IssuingDemo
{
    static public class Program
    {
        static async Task Main(string[] args)
        {
            PanelMaking iss = new PanelMaking();
            await iss.GeneratePanelMaking();
        }

/*        public static async Task GeneratePanelMaking()
        {
            var intSq = new List<PanelModel>();
            var intAng = new List<PanelModel>();
            var extSq = new List<PanelModel>();
            var extAng = new List<PanelModel>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var reader = new StreamReader(@"C:\Users\mateusz.konopka\Work Folders\Desktop\Issuing 2.0\19007GF_elevations.csv"))
            {
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    csv.Read();
                    while (csv.Read())
                    {
                        var record = new PanelModel
                        {
                            PanelType = csv.GetField<string>(0),
                            PanelRef = csv.GetField<string>(1),
                            PanelSquareAngled = csv.GetField<string>(2),
                            Length = csv.GetField<double>(3),
                            Height = csv.GetField<double>(4),
                            Area = csv.GetField<double>(5),
                            Weight = csv.GetField<double>(6),
                            Qty = csv.GetField<int>(7)
                        };

                        if (record.PanelType == "Int" && record.PanelSquareAngled == "Sq") intSq.Add(record);
                        if (record.PanelType == "Int" && record.PanelSquareAngled == "Ang") intAng.Add(record);
                        if (record.PanelType == "Ext" && record.PanelSquareAngled == "Sq") extSq.Add(record);
                        if (record.PanelType == "Ext" && record.PanelSquareAngled == "Ang") extAng.Add(record);

                    }
                    //
                    var file = new FileInfo(@"C:\Users\mateusz.konopka\Work Folders\Desktop\Issuing 2.0\Book1.xlsx");
                    DeleteIfExist(file);

                    if (intSq.Count > 0)
                    {
                        await SaveExcelFile(intSq, file, "INT SQ Panel");
                    }

                    if (intAng.Count > 0)
                    {
                        await SaveExcelFile(intAng, file, "INT ANG Panel");
                    }

                    if (extSq.Count > 0)
                    {
                        await SaveExcelFile(extSq, file, "EXT SQ Panel");
                    }

                    if (extAng.Count > 0)
                    {
                        await SaveExcelFile(extAng, file, "EXT ANG Panel");
                    }

                }
            }
        }

        private static async Task SaveExcelFile(IEnumerable<PanelModel> panels, FileInfo file, string wsName)
        {
            var list = panels
                   .Select(x => new { x.PanelRef, x.Length, x.Height, x.Area, x.Weight, x.Qty })
                   .ToList();

            using (var package = new ExcelPackage(file))
            {
                var ws = package.Workbook.Worksheets.Add(wsName);
                
                var wsAreaSum = panels.Sum(x => x.Area);
                var wsLengthSum = panels.Sum(x => x.Length);
                var wsQtySum = panels.Sum(x => x.Qty);

                CreateTemplateTop(ws);
                CreateSummaryMakePanels(ws, wsAreaSum, wsLengthSum, wsQtySum);
                CreateHeadersPanelMaking(ws);

                var range = ws.Cells["A14"].LoadFromCollection(list, false);
                await package.SaveAsync();
            }

        }

        private static void CreateHeadersPanelMaking(ExcelWorksheet ws)
        {
            ws.Cells["A13"].Value = "Panel Ref";
            ws.Cells["A13"].Style.Font.Bold = true;
            ws.Cells["A13"].Style.Font.Italic = true;

            ws.Cells["B13"].Value = "Length (mm)";
            ws.Cells["B13"].Style.Font.Bold = true;
            ws.Cells["B13"].Style.Font.Italic = true;

            ws.Cells["C13"].Value = "Height (mm)";
            ws.Cells["C13"].Style.Font.Bold = true;
            ws.Cells["C13"].Style.Font.Italic = true;

            ws.Cells["D13"].Value = "Area (m2)";
            ws.Cells["D13"].Style.Font.Bold = true;
            ws.Cells["D13"].Style.Font.Italic = true;

            ws.Cells["E13"].Value = "Weight (kg)";
            ws.Cells["E13"].Style.Font.Bold = true;
            ws.Cells["E13"].Style.Font.Italic = true;

            ws.Cells["F13"].Value = "Qty";
            ws.Cells["F13"].Style.Font.Bold = true;
            ws.Cells["F13"].Style.Font.Italic = true;

            ws.Cells["G13"].Value = "Remarks";
            ws.Cells["G13"].Style.Font.Bold = true;
            ws.Cells["G13"].Style.Font.Italic = true;
        }

        private static void CreateSummaryMakePanels(ExcelWorksheet ws, double wsAreaSum, double wsLengthSum, int qtySum)
        {
            ws.Cells["A10"].Value = "SUMMARY";
            ws.Cells["A10"].Style.Font.Bold = true;
            ws.Cells["A10"].Style.Font.Italic = true;
            ws.Cells["A11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            ws.Cells["B10"].Value = "Length (m)";
            ws.Cells["B10"].Style.Font.Bold = true;
            ws.Cells["B10"].Style.Font.Italic = true;
            ws.Cells["B11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            ws.Cells["C10"].Value = "Area (m2)";
            ws.Cells["C10"].Style.Font.Bold = true;
            ws.Cells["C10"].Style.Font.Italic = true;
            ws.Cells["C11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            ws.Cells["D10"].Value = "Qty";
            ws.Cells["D10"].Style.Font.Bold = true;
            ws.Cells["D10"].Style.Font.Italic = true;
            ws.Cells["D11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            ws.Cells["A10:G11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            ws.Cells["D11"].Value = qtySum;
            ws.Cells["C11"].Value = wsAreaSum;
            ws.Cells["B11"].Value = Math.Round(wsLengthSum * 0.001, 1);
        }

        private static void CreateTemplateTop(ExcelWorksheet ws)
        {
            ws.Column(1).Width = 16.29 * 1.0463;
            ws.Column(2).Width = 12.29 * 1.062;
            ws.Column(3).Width = 12.29 * 1.062;
            ws.Column(4).Width = 10.29 * 1.0463;
            ws.Column(5).Width = 12.29 * 1.062;
            ws.Column(6).Width = 4 * 1.216;
            ws.Column(7).Width = 14.29 * 1.062;

            ws.Cells["A1"].Value = "KINGSPAN";
            ws.Cells["A1:G1"].Merge = true;
            ws.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["A1"].Style.Font.Size = 18;
            ws.Row(1).Height = 23.25;
            ws.Cells["A1:G1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            ws.Cells["A2:G7"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            ws.Cells["A2"].Value = "CLIENT";
            ws.Cells["A3"].Value = "JOB NO.";
            ws.Cells["A4"].Value = "PLOT";
            ws.Cells["A5"].Value = "SET REF";
            ws.Cells["A6"].Value = "CREATED BY";
            ws.Cells["A7"].Value = "DATE";

            ws.Cells["B2"].Value = "ClientName";
            ws.Cells["B2:D2"].Merge = true;
            ws.Cells["B3"].Value = "25-335";
            ws.Cells["B3:D3"].Merge = true;
            ws.Cells["B4"].Value = "1";
            ws.Cells["B4:D4"].Merge = true;
            ws.Cells["B5"].Value = "GF Walls";
            ws.Cells["B5:D5"].Merge = true;
            ws.Cells["B6"].Value = "Kingspan";
            ws.Cells["B6:D6"].Merge = true;
            ws.Cells["B7"].Value = DateTime.Now;
            ws.Cells["B7"].Style.Numberformat.Format = "dd-mmm-yyyy";
            ws.Cells["B7:D7"].Merge = true;
            ws.Cells["B7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            ws.Cells["E2:G7"].Merge = true;

            Image img = Image.FromFile(@"C:\Users\mateusz.konopka\Work Folders\Desktop\Issuing 2.0\newlogo2.png");
            ExcelPicture pic = ws.Drawings.AddPicture("KingspanLogo", img);
            pic.LockAspectRatio = false;
            pic.SetPosition(1, 0, 4, 0);
            pic.SetSize(19);

            ws.Cells["A8"].Value = ws.Name + " List";
            ws.Cells["A8:G8"].Merge = true;
            ws.Cells["A8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["A8"].Style.Font.Size = 16;
            ws.Row(1).Height = 21;
            ws.Cells["A8:G8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }

        private static void DeleteIfExist(FileInfo file)
        {
            if (file.Exists) file.Delete();
        }*/

    }
}
