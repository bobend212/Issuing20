using CsvHelper;
using CsvHelper.Configuration;
using IssuingDemoLogger;
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
    public class PanelMaking : Templates
    {
        public async Task GeneratePanelMaking()
        {
            var intSq = new List<PanelMakingModel>();
            var intAng = new List<PanelMakingModel>();
            var extSq = new List<PanelMakingModel>();
            var extAng = new List<PanelMakingModel>();

            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = true,
            };

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //using (var reader = new StreamReader(@"C:\Users\mateusz.konopka\Work Folders\Desktop\Issuing 2.0\19007GF_elevations.csv"))
            using (var reader = new StreamReader($@"C:\MiTek\UK\jobs\{_mbaJob}\Attachments\{_mbaJob}_elevations.csv"))
            {
                using (var csv = new CsvReader(reader, config))
                {
                    csv.Read();
                    csv.ReadHeader();
                    while (csv.Read())
                    {
                        var record = new PanelMakingModel
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
                    var file = new FileInfo($@"C:\MiTek\UK\jobs\{_mbaJob}\Attachments\{_jobNo}_issuing.xlsx");
                    //DeleteIfExist(file);

                    if (intSq.Count > 0)
                    {
                        await SaveExcelFile(intSq, file, "INT SQ Panel");
                    }

                    if (extSq.Count > 0)
                    {
                        await SaveExcelFile(extSq, file, "EXT SQ Panel");
                    }

                    if (intAng.Count > 0)
                    {
                        await SaveExcelFile(intAng, file, "INT ANG Panel");
                    }

                    if (extAng.Count > 0)
                    {
                        await SaveExcelFile(extAng, file, "EXT ANG Panel");
                    }

                }
            }
        }

        private async Task SaveExcelFile(IEnumerable<PanelMakingModel> panels, FileInfo file, string wsName)
        {

            var list = panels
                   .Select(x => new { x.PanelRef, x.Length, x.Height, x.Area, x.Weight, x.Qty })
                   .OrderBy(x => x.Area)
                   .ToList();

                var wsAreaSum = panels.Sum(x => x.Area);
                var wsLengthSum = panels.Sum(x => x.Length);
                var wsQtySum = panels.Sum(x => x.Qty);

                await AddTS(file, "TS." + wsName, wsAreaSum);

            using (var package = new ExcelPackage(file))
                {
                    var ws = package.Workbook.Worksheets.Add(wsName);

                    CreateTemplateTop(ws);
                    AddSummaryFieldsPanelMaking(ws, wsAreaSum, wsLengthSum, wsQtySum);
                    AddHeadersPanelMaking(ws);

                    var range = ws.Cells["A14"].LoadFromCollection(list, false);
                    await package.SaveAsync();
                }

        }

        private void AddHeadersPanelMaking(ExcelWorksheet ws)
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

        private void AddSummaryFieldsPanelMaking(ExcelWorksheet ws, double wsAreaSum, double wsLengthSum, int qtySum)
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


    }
}
