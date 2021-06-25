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
    public class PanelCutting : Templates
    {
        public async Task GeneratePanelCutting()
        {
            var intSq = new List<PanelCuttingModel>();
            var intAng = new List<PanelCuttingModel>();
            var extSq = new List<PanelCuttingModel>();
            var extAng = new List<PanelCuttingModel>();
            var ultSq = new List<PanelCuttingModel>();
            var ultAng = new List<PanelCuttingModel>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var reader = new StreamReader(@"C:\Users\mateusz.konopka\Work Folders\Desktop\Issuing 2.0\19007GF_cutting.csv"))
            {
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    csv.Read();
                    csv.ReadHeader();
                    while (csv.Read())
                    {
                        var record = new PanelCuttingModel
                        {
                            PanelType = csv.GetField<string>(0),
                            PanelRef = csv.GetField<string>(1),
                            PanelSquareAngled = csv.GetField<string>(2),
                            Material = csv.GetField<string>(3),
                            Grade = csv.GetField<string>(4),
                            Width = csv.GetField<double>(5),
                            Height = csv.GetField<double>(6),
                            Length = csv.GetField<double>(7),
                            LeftCut = csv.GetField<double>(8),
                            RightCut = csv.GetField<double>(9),
                            Qty = csv.GetField<int>(10),
                        };

                        if (record.PanelType == "Int" && record.PanelSquareAngled == "Sq") intSq.Add(record);
                        if (record.PanelType == "Int" && record.PanelSquareAngled == "Ang") intAng.Add(record);
                        if (record.PanelType == "Ext" && record.PanelSquareAngled == "Sq" && record.Material != "Batten") extSq.Add(record);
                        if (record.PanelType == "Ext" && record.PanelSquareAngled == "Ang" && record.Material != "Batten") extAng.Add(record);
                        if (record.PanelType == "Ext" && record.PanelSquareAngled == "Sq" && record.Material == "Batten") ultSq.Add(record);
                        if (record.PanelType == "Ext" && record.PanelSquareAngled == "Ang" && record.Material == "Batten") ultAng.Add(record);
                    }

                    //
                    var file = new FileInfo(@"C:\Users\mateusz.konopka\Work Folders\Desktop\Issuing 2.0\"
                        + _jobNo + "_issuing.xlsx");
                    //DeleteIfExist(file);

                    if (ultSq.Count > 0)
                    {
                        await SaveExcelFile(ultSq, file, "SQ Bat");
                    }

                    if (ultAng.Count > 0)
                    {
                        await SaveExcelFile(ultAng, file, "ANG Bat");
                    }

                    if (intSq.Count > 0)
                    {
                        await SaveExcelFile(intSq, file, "INT SQ Cut");
                    }

                    if (extSq.Count > 0)
                    {
                        await SaveExcelFile(extSq, file, "EXT SQ Cut");
                    }

                    if (intAng.Count > 0)
                    {
                        await SaveExcelFile(intAng, file, "INT ANG Cut");
                    }

                    if (extAng.Count > 0)
                    {
                        await SaveExcelFile(extAng, file, "EXT ANG Cut");
                    }

                }
            }
        }

        private async Task SaveExcelFile(IEnumerable<PanelCuttingModel> panels, FileInfo file, string wsName)
        {
            var panelRefs = panels.Select(x => new { x.PanelRef }).Distinct().ToList();

            var panelTimbers = panels
               .Select(x => new { x.PanelRef, x.Material, x.Grade, x.Width, x.Height, x.Length, x.LeftCut, x.RightCut, x.Qty })
               .OrderByDescending(x => x.Length)
               .ToList();

            var bulkStuds = panels
                .GroupBy(x => new { x.Width, x.Height, x.Length})
                .Select(group => new
                {
                    PanelRef = group.Select(x => x.PanelRef),
                    Material = group.Select(x => x.Material),
                    Grade = group.Select(x => x.Grade),
                    Width = group.Key.Width,
                    Height = group.Key.Height,
                    Length = group.Key.Length,
                    LeftCut = group.Select(x => x.LeftCut),
                    RightCut = group.Select(x => x.RightCut),
                    Total = group.Sum(x => x.Qty),
                })
                    .Where(x => x.Total > 10)
                    .OrderByDescending(x => x.Length)
                    .ToList();

            foreach (var item in bulkStuds)
            {
                panelTimbers.RemoveAll(x => x.Length == item.Length && x.Width == item.Width && x.Height == item.Height);
            }

            foreach (var item in panelTimbers)
            {
                panelRefs.RemoveAll(x => !panelTimbers.Select(y => y.PanelRef).Contains(x.PanelRef));
            }

            var totalLength = panelTimbers.Sum(x => x.Length);
            var totalQty = panelTimbers.Sum(x => x.Qty);

            await AddTS(file, "TS." + wsName, 0, totalLength, totalQty);
            using (var package = new ExcelPackage(file))
            {
                var ws = package.Workbook.Worksheets.Add(wsName);

                CreateTemplateTop(ws);
                PrettifyCuttingTemplate(ws);

                var maxRow = ws.Cells
                    .Select(c => c.Start.Row)
                    .Max();
                maxRow++;

                if (panelTimbers.Count > 0)
                {
                    for (int i = 0; i < panelRefs.Count; i++)
                    {
                        maxRow++;
                        ws.Cells[maxRow, 1].Value = panelRefs[i].PanelRef;
                        ws.Cells[maxRow, 1].Style.Font.Bold = true;
                        maxRow++;

                        AddHeadersPanelCutting(ws, maxRow);

                        for (int j = 0; j < panelTimbers.Count; j++)
                        {
                            if (panelRefs[i].PanelRef == panelTimbers[j].PanelRef)
                            {
                                maxRow++;

                                ws.Cells[maxRow, 1].Value = panelTimbers[j].Material;
                                ws.Cells[maxRow, 2].Value = panelTimbers[j].Grade;
                                ws.Cells[maxRow, 3].Value = panelTimbers[j].Width;
                                ws.Cells[maxRow, 4].Value = panelTimbers[j].Height;
                                ws.Cells[maxRow, 5].Value = panelTimbers[j].Length;
                                ws.Cells[maxRow, 6].Value = panelTimbers[j].LeftCut;
                                ws.Cells[maxRow, 7].Value = panelTimbers[j].RightCut;
                                ws.Cells[maxRow, 8].Value = panelTimbers[j].Qty;

                                AllignLeft(ws, maxRow, 1);
                                AllignLeft(ws, maxRow, 2);
                                AllignLeft(ws, maxRow, 3);
                                AllignLeft(ws, maxRow, 4);
                                AllignLeft(ws, maxRow, 5);
                                AllignLeft(ws, maxRow, 6);
                                AllignLeft(ws, maxRow, 7);
                                AllignLeft(ws, maxRow, 8);
                            }
                        }
                        maxRow++;
                    }
                }

                if (bulkStuds.Count > 0)
                {
                    var bulkStudsWorksheet = package.Workbook.Worksheets.Add(ws.Name.Replace("Cut", "Bulks"));
                    
                    await AddTS(file, "TS." + bulkStudsWorksheet, 2);
                    
                    CreateTemplateTop(bulkStudsWorksheet);
                    PrettifyCuttingTemplate(bulkStudsWorksheet);
                    AddHeadersPanelCutting(bulkStudsWorksheet, 10);

                    bulkStudsWorksheet.Cells["A8"].Value = ws.Name.Replace("Cut", "Bulk Studs");

                    var maxRowBulks = bulkStudsWorksheet.Cells.Select(c => c.Start.Row).Max();

                    for (int i = 0; i < bulkStuds.Count; i++)
                    {
                        maxRowBulks++;
                        bulkStudsWorksheet.Cells[maxRowBulks, 1].Value = bulkStuds[i].Material;
                        bulkStudsWorksheet.Cells[maxRowBulks, 2].Value = bulkStuds[i].Grade;
                        bulkStudsWorksheet.Cells[maxRowBulks, 3].Value = bulkStuds[i].Width;
                        bulkStudsWorksheet.Cells[maxRowBulks, 4].Value = bulkStuds[i].Height;
                        bulkStudsWorksheet.Cells[maxRowBulks, 5].Value = bulkStuds[i].Length;
                        bulkStudsWorksheet.Cells[maxRowBulks, 6].Value = bulkStuds[i].LeftCut;
                        bulkStudsWorksheet.Cells[maxRowBulks, 7].Value = bulkStuds[i].RightCut;
                        bulkStudsWorksheet.Cells[maxRowBulks, 8].Value = bulkStuds[i].Total;

                        AllignLeft(bulkStudsWorksheet, maxRowBulks, 1);
                        AllignLeft(bulkStudsWorksheet, maxRowBulks, 2);
                        AllignLeft(bulkStudsWorksheet, maxRowBulks, 3);
                        AllignLeft(bulkStudsWorksheet, maxRowBulks, 4);
                        AllignLeft(bulkStudsWorksheet, maxRowBulks, 5);
                        AllignLeft(bulkStudsWorksheet, maxRowBulks, 6);
                        AllignLeft(bulkStudsWorksheet, maxRowBulks, 7);
                        AllignLeft(bulkStudsWorksheet, maxRowBulks, 8);
                    }
                }
                await package.SaveAsync();
            }

        }

        private static void PrettifyCuttingTemplate(ExcelWorksheet ws)
        {
            ws.InsertColumn(3, 1);
            ws.InsertColumn(7, 1);
            ws.Cells["A1:I1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            ws.Cells["A8:I8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            ws.Cells["B1:G1"].AutoFitColumns();
            ws.Cells["I1"].AutoFitColumns();
        }

        private static void AllignLeft(ExcelWorksheet ws, int maxRow, int cell)
        {
            ws.Cells[maxRow, cell].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        }

        private static void AddHeadersPanelCutting(ExcelWorksheet ws, int maxRow)
        {
            ws.Cells[maxRow, 1].Value = "Description";
            ws.Cells[maxRow, 1].Style.Font.Bold = true;
            ws.Cells[maxRow, 1].Style.Font.Italic = true;

            ws.Cells[maxRow, 2].Value = "Material";
            ws.Cells[maxRow, 2].Style.Font.Bold = true;
            ws.Cells[maxRow, 2].Style.Font.Italic = true;

            ws.Cells[maxRow, 3].Value = "Width";
            ws.Cells[maxRow, 3].Style.Font.Bold = true;
            ws.Cells[maxRow, 3].Style.Font.Italic = true;

            ws.Cells[maxRow, 4].Value = "Height";
            ws.Cells[maxRow, 4].Style.Font.Bold = true;
            ws.Cells[maxRow, 4].Style.Font.Italic = true;

            ws.Cells[maxRow, 5].Value = "Length";
            ws.Cells[maxRow, 5].Style.Font.Bold = true;
            ws.Cells[maxRow, 5].Style.Font.Italic = true;

            ws.Cells[maxRow, 6].Value = "Left Cut";
            ws.Cells[maxRow, 6].Style.Font.Bold = true;
            ws.Cells[maxRow, 6].Style.Font.Italic = true;

            ws.Cells[maxRow, 7].Value = "Right Cut";
            ws.Cells[maxRow, 7].Style.Font.Bold = true;
            ws.Cells[maxRow, 7].Style.Font.Italic = true;

            ws.Cells[maxRow, 8].Value = "Qty";
            ws.Cells[maxRow, 8].Style.Font.Bold = true;
            ws.Cells[maxRow, 8].Style.Font.Italic = true;

            ws.Cells[maxRow, 9].Value = "Remarks";
            ws.Cells[maxRow, 9].Style.Font.Bold = true;
            ws.Cells[maxRow, 9].Style.Font.Italic = true;
        }

    }
}
