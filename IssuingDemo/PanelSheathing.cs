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
    public class PanelSheathing : Templates
    {
        public async Task GeneratePanelSheathing()
        {
            var intSq = new List<PanelSheathingModel>();
            var intAng = new List<PanelSheathingModel>();
            var extSq = new List<PanelSheathingModel>();
            var extAng = new List<PanelSheathingModel>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var reader = new StreamReader(@"C:\Users\mateusz.konopka\Work Folders\Desktop\Issuing 2.0\19007GF_sheathing.csv"))
            {
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    csv.Read();
                    csv.ReadHeader();
                    while (csv.Read())
                    {
                        var record = new PanelSheathingModel
                        {
                            PanelType = csv.GetField<string>(0),
                            PanelRef = csv.GetField<string>(1),
                            PanelSquareAngled = csv.GetField<string>(2),
                            Type = csv.GetField<string>(3),
                            Material = csv.GetField<string>(4),
                            Thickness = csv.GetField<double>(5),
                            Height = csv.GetField<double>(6),
                            Width = csv.GetField<double>(7),
                            Qty = csv.GetField<int>(8)
                        };

                        if (record.PanelType == "Int" && record.PanelSquareAngled == "Sq") intSq.Add(record);
                        if (record.PanelType == "Int" && record.PanelSquareAngled == "Ang") intAng.Add(record);
                        if (record.PanelType == "Ext" && record.PanelSquareAngled == "Sq") extSq.Add(record);
                        if (record.PanelType == "Ext" && record.PanelSquareAngled == "Ang") extAng.Add(record);

                    }
                    //
                    var file = new FileInfo(@"C:\Users\mateusz.konopka\Work Folders\Desktop\Issuing 2.0\"
                        + _jobNo + "_issuing.xlsx");
                    //DeleteIfExist(file);

                    if (intSq.Count > 0)
                    {
                        await SaveExcelFile(intSq, file, "INT SQ Sh");
                    }

                    if (extSq.Count > 0)
                    {
                        await SaveExcelFile(extSq, file, "EXT SQ Sh");
                    }

                    if (intAng.Count > 0)
                    {
                        await SaveExcelFile(intAng, file, "INT ANG Sh");
                    }

                    if (extAng.Count > 0)
                    {
                        await SaveExcelFile(extAng, file, "EXT ANG Sh");
                    }

                }
            }
        }
        
        private async Task SaveExcelFile(IEnumerable<PanelSheathingModel> panels, FileInfo file, string wsName)
        {

            var panelRefs = panels
               .Select(x => new { x.PanelRef })
               .Distinct()
               .ToList();

            var panelSheathing = panels
               .Select(x => new { x.PanelRef, x.Type, x.Material, x.Thickness, x.Height, x.Width, x.Qty })
               .OrderByDescending(x => x.Thickness)
               .ToList();

            double areaSum = 0.0;
            double qtySum = 0;

            foreach (var item in panelSheathing)
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
                ws.Column(2).Width = 16.29 * 1.0463;

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

                    AddHeadersPanelSheathing(ws, maxRow);

                    for (int j = 0; j < panelSheathing.Count; j++)
                    {
                        if (panelRefs[i].PanelRef == panelSheathing[j].PanelRef)
                        {
                            maxRow++;

                            ws.Cells[maxRow, 1].Value = panelSheathing[j].Type;
                            ws.Cells[maxRow, 2].Value = panelSheathing[j].Material;
                            ws.Cells[maxRow, 3].Value = panelSheathing[j].Thickness;
                            ws.Cells[maxRow, 4].Value = panelSheathing[j].Height;
                            ws.Cells[maxRow, 5].Value = panelSheathing[j].Width;
                            ws.Cells[maxRow, 6].Value = panelSheathing[j].Qty;

                            AllignLeft(ws, maxRow, 1);
                            AllignLeft(ws, maxRow, 2);
                            AllignLeft(ws, maxRow, 3);
                            AllignLeft(ws, maxRow, 4);
                            AllignLeft(ws, maxRow, 5);
                            AllignLeft(ws, maxRow, 6);
                        }
                    }
                    maxRow++;
                }
                await package.SaveAsync();
            }

        }

        private static void AllignLeft(ExcelWorksheet ws, int maxRow, int cell)
        {
            ws.Cells[maxRow, cell].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        }

        private static void AddHeadersPanelSheathing(ExcelWorksheet ws, int maxRow)
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
        }
    }
}
