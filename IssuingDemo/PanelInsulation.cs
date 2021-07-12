using CsvHelper;
using IssuingDemo.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IssuingDemo
{
    public class PanelInsulation : Templates
    {
        public async Task GeneratePanelInsulation()
        {
            var intSq = new List<PanelInsulationModel>();
            var intAng = new List<PanelInsulationModel>();
            var extSq = new List<PanelInsulationModel>();
            var extAng = new List<PanelInsulationModel>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var reader = new StreamReader($@"C:\MiTek\UK\jobs\{_mbaJob}\Attachments\{_mbaJob}_insulation.csv"))
            {
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    csv.Read();
                    csv.ReadHeader();
                    while (csv.Read())
                    {
                        var record = new PanelInsulationModel
                        {
                            PanelType = csv.GetField<string>(0),
                            PanelRef = csv.GetField<string>(1),
                            PanelSquareAngled = csv.GetField<string>(2),
                            Number = csv.GetField<int>(3),
                            Material = csv.GetField<string>(4),
                            Thickness = csv.GetField<double>(5),
                            Height = csv.GetField<double>(6),
                            Width = csv.GetField<double>(7),
                            Qty = csv.GetField<int>(8),
                            Area = csv.GetField<double>(9)
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
                        await SaveExcelFile(intSq, file, "INT SQ Ins");
                    }

                    if (extSq.Count > 0)
                    {
                        await SaveExcelFile(extSq, file, "EXT SQ Ins");
                    }

                    if (intAng.Count > 0)
                    {
                        await SaveExcelFile(intAng, file, "INT ANG Ins");
                    }

                    if (extAng.Count > 0)
                    {
                        await SaveExcelFile(extAng, file, "EXT ANG Ins");
                    }

                }
            }
        }

        private async Task SaveExcelFile(IEnumerable<PanelInsulationModel> panels, FileInfo file, string wsName)
        {

            var panelRefs = panels
               .Select(x => new { x.PanelRef })
               .Distinct()
               .ToList();

            var panelInsulation = panels
               .Select(x => new { x.PanelRef, x.Number, x.Material, x.Thickness, x.Height, x.Width, x.Qty, x.Area })
               .OrderBy(x => x.Number)
               .ToList();

            double areaSum = 0.0;
            double qtySum = 0;

            foreach (var item in panelInsulation)
            {
                areaSum += item.Height * item.Width * item.Qty;
                if (Math.Abs(item.Height - 2400) > 4 && Math.Abs(item.Height - 1200) > 4) qtySum += item.Qty;
                if (Math.Abs(item.Width - 2400) > 4 && Math.Abs(item.Width - 1200) > 4) qtySum += item.Qty;
            }


            await AddTS(file, "TS." + wsName, areaSum, 0, qtySum);
            using (var package = new ExcelPackage(file))
            {
                var ws = package.Workbook.Worksheets.Add(wsName);

                CreateTemplateTop(ws);
                ws.Cells["C1:D1"].AutoFitColumns();
                ws.Column(6).Width = 10 * 1.0463;
                ws.Column(7).Width = 9 * 1.0463;

                var cells = ws.Cells;

                var maxRow = cells
                    .Select(c => c.Start.Row)
                    .Max();
                maxRow++;

                for (int i = 0; i < panelRefs.Count; i++)
                {
                    maxRow++;
                    ws.Cells[maxRow, 1].Value = panelRefs[i].PanelRef;
                    ws.Cells[maxRow, 1].Style.Font.Bold = true;

                    maxRow++;

                    AddHeadersPanelInsulation(ws, maxRow);
                    double totalArea = 0;
                    for (int j = 0; j < panelInsulation.Count; j++)
                    {
                        if (panelRefs[i].PanelRef == panelInsulation[j].PanelRef)
                        {
                            maxRow++;

                            ws.Cells[maxRow, 1].Value = panelInsulation[j].Number;
                            ws.Cells[maxRow, 2].Value = panelInsulation[j].Material;
                            ws.Cells[maxRow, 3].Value = panelInsulation[j].Thickness;
                            ws.Cells[maxRow, 4].Value = panelInsulation[j].Height;
                            ws.Cells[maxRow, 5].Value = panelInsulation[j].Width;
                            ws.Cells[maxRow, 6].Value = panelInsulation[j].Qty;
                            ws.Cells[maxRow, 7].Value = panelInsulation[j].Area;

                            AllignLeft(ws, maxRow, 1);
                            AllignLeft(ws, maxRow, 2);
                            AllignLeft(ws, maxRow, 3);
                            AllignLeft(ws, maxRow, 4);
                            AllignLeft(ws, maxRow, 5);
                            AllignLeft(ws, maxRow, 6);
                            AllignLeft(ws, maxRow, 7);

                            totalArea += panelInsulation[j].Area;
                        }

                    }
                    maxRow++;
                    ws.Cells[maxRow, 6].Value = "Total Area:";
                    ws.Cells[maxRow, 7].Value = Math.Round(totalArea, 2) + "m2";

                    ws.Cells[maxRow, 6].Style.Font.Bold = true;
                    ws.Cells[maxRow, 7].Style.Font.Bold = true;
                    AllignLeft(ws, maxRow, 6);
                    AllignLeft(ws, maxRow, 7);
                    maxRow++;
                }
                await package.SaveAsync();
            }

        }

        private static void AllignLeft(ExcelWorksheet ws, int maxRow, int cell)
        {
            ws.Cells[maxRow, cell].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        }

        private static void AddHeadersPanelInsulation(ExcelWorksheet ws, int maxRow)
        {
            ws.Cells[maxRow, 1].Value = "Description";
            ws.Cells[maxRow, 1].Style.Font.Bold = true;
            ws.Cells[maxRow, 1].Style.Font.Italic = true;

            ws.Cells[maxRow, 2].Value = "Material";
            ws.Cells[maxRow, 2].Style.Font.Bold = true;
            ws.Cells[maxRow, 2].Style.Font.Italic = true;

            ws.Cells[maxRow, 3].Value = "Thickness";
            ws.Cells[maxRow, 3].Style.Font.Bold = true;
            ws.Cells[maxRow, 3].Style.Font.Italic = true;

            ws.Cells[maxRow, 4].Value = "Height";
            ws.Cells[maxRow, 4].Style.Font.Bold = true;
            ws.Cells[maxRow, 4].Style.Font.Italic = true;

            ws.Cells[maxRow, 5].Value = "Width";
            ws.Cells[maxRow, 5].Style.Font.Bold = true;
            ws.Cells[maxRow, 5].Style.Font.Italic = true;

            ws.Cells[maxRow, 6].Value = "Qty";
            ws.Cells[maxRow, 6].Style.Font.Bold = true;
            ws.Cells[maxRow, 6].Style.Font.Italic = true;

            ws.Cells[maxRow, 7].Value = "Area";
            ws.Cells[maxRow, 7].Style.Font.Bold = true;
            ws.Cells[maxRow, 7].Style.Font.Italic = true;
        }
    }
}
