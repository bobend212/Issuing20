using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;

namespace IssuingDemo
{
    public class Templates
    {
        public string _clientName;
        public string _jobNo;
        public string _plot;
        public string _setRef;
        public string _mbaJob;

        public void CreateTemplateTop(ExcelWorksheet ws)
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

            ws.Cells["B2"].Value = _clientName;
            ws.Cells["B2:D2"].Merge = true;
            ws.Cells["B3"].Value = _jobNo;
            ws.Cells["B3:D3"].Merge = true;
            ws.Cells["B4"].Value = _plot;
            ws.Cells["B4:D4"].Merge = true;
            ws.Cells["B5"].Value = _setRef;
            ws.Cells["B5:D5"].Merge = true;
            ws.Cells["B6"].Value = "Kingspan";
            ws.Cells["B6:D6"].Merge = true;
            ws.Cells["B7"].Value = DateTime.Now;
            ws.Cells["B7"].Style.Numberformat.Format = "dd-mmm-yyyy";
            ws.Cells["B7:D7"].Merge = true;
            ws.Cells["B7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            ws.Cells["E2:G7"].Merge = true;

            InsertLogo(ws);

            if (ws.Name.Contains("Panel")) ws.Cells["A8"].Value = ws.Name + " List";
            if (ws.Name.Contains("Cut")) ws.Cells["A8"].Value = ws.Name + "ting";
            if (ws.Name.Contains("Bat")) ws.Cells["A8"].Value = "Ultima Inside " + ws.Name + "ten List";
            if (ws.Name.Contains("Sh")) ws.Cells["A8"].Value = ws.Name + "eathing";


            ws.Cells["A8:G8"].Merge = true;
            ws.Cells["A8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["A8"].Style.Font.Size = 16;
            ws.Cells["A8:G8"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }

        private static void InsertLogo(ExcelWorksheet ws)
        {
            Image img = Image.FromFile(@"C:\Users\mateusz.konopka\Work Folders\Desktop\Issuing 2.0\newlogo2.png");
            ExcelPicture pic = ws.Drawings.AddPicture("KingspanLogo", img);
            pic.LockAspectRatio = false;
            pic.SetPosition(1, 0, 4, 0);
            pic.SetSize(19);
        }

        public void DeleteIfExist(FileInfo file)
        {
            if (file.Exists) file.Delete();
        }

        //Top Sheets
        public async Task AddTS(FileInfo file, string wsName, [Optional] double areaSum, [Optional] double lengthSum,[Optional] double qtySum)
        {
            using (var package = new ExcelPackage(file))
            {
                var topSheet = package.Workbook.Worksheets.Add(wsName);

                DetermineTopSheetColorTab(topSheet);

                //top headers
                topSheet.Cells["A1"].Value = _clientName.ToUpper();
                topSheet.Cells["A1:I2"].Merge = true;
                topSheet.Cells["A1"].Style.Font.Bold = true;
                topSheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                topSheet.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                topSheet.Cells["A1"].Style.Font.Size = 22;

                topSheet.Cells["A3"].Value = "PLOTS: " + _plot;
                topSheet.Cells["A3:I4"].Merge = true;
                topSheet.Cells["A3"].Style.Font.Bold = true;
                topSheet.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                topSheet.Cells["A3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                topSheet.Cells["A3"].Style.Font.Size = 22;

                topSheet.Cells["A5"].Value = "AREA: " + _setRef.ToUpper();
                topSheet.Cells["A5:I6"].Merge = true;
                topSheet.Cells["A5"].Style.Font.Bold = true;
                topSheet.Cells["A5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                topSheet.Cells["A5"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                topSheet.Cells["A5"].Style.Font.Size = 22;

                //info
                topSheet.Cells["A8"].Value = "Delivery Date:";
                topSheet.Cells["A8"].Style.Font.Bold = true;
                topSheet.Cells["A8"].Style.Font.Size = 20;

                topSheet.Cells["A10"].Value = "Treatment: " + Treatment(wsName);
                topSheet.Cells["A10"].Style.Font.Bold = true;
                topSheet.Cells["A10"].Style.Font.Size = 20;

                topSheet.Cells["A12"].Value = "Paper:";
                topSheet.Cells["A12"].Style.Font.Bold = true;
                topSheet.Cells["A12"].Style.Font.Size = 20;

                topSheet.Cells["A14"].Value = "Timesheet Code:";
                topSheet.Cells["A14"].Style.Font.Bold = true;
                topSheet.Cells["A14"].Style.Font.Size = 20;

                topSheet.Cells["A16"].Value = "Item: " + ItemName(wsName);
                topSheet.Cells["A16"].Style.Font.Bold = true;
                topSheet.Cells["A16"].Style.Font.Size = 20;

                Output(topSheet, areaSum, lengthSum, qtySum);
                topSheet.Cells["A19"].Style.Font.Bold = true;
                topSheet.Cells["A19"].Style.Font.Size = 20;

                topSheet.Cells["A21"].Value = "Job Number: " + _jobNo;
                topSheet.Cells["A21"].Style.Font.Bold = true;
                topSheet.Cells["A21"].Style.Font.Size = 20;

                topSheet.Cells["A23:I23"].Style.Border.Top.Style = ExcelBorderStyle.Thick;

                topSheet.Cells["A25"].Value = "Name:";
                topSheet.Cells["A25"].Style.Font.Bold = true;
                topSheet.Cells["A25"].Style.Font.Size = 20;

                topSheet.Cells["A27"].Value = "Time Started:";
                topSheet.Cells["A27"].Style.Font.Bold = true;
                topSheet.Cells["A27"].Style.Font.Size = 20;

                topSheet.Cells["A29"].Value = "Time Complete:";
                topSheet.Cells["A29"].Style.Font.Bold = true;
                topSheet.Cells["A29"].Style.Font.Size = 20;

                await package.SaveAsync();
            }
        }

        private static void Output(ExcelWorksheet topSheet, double areaSum, double lengthSum, double qtySum)
        {
            if(topSheet.Name.Contains("Panel"))  topSheet.Cells["A19"].Value = "Output: " + Math.Round(areaSum, 0) + " m2";

            if (topSheet.Name.Contains("Cut") || topSheet.Name.Contains("Bat")) topSheet.Cells["A19"].Value = "Output: " 
                    + qtySum + " CUTS, " + Math.Round(lengthSum * 0.001, 0) + " ml";

            if (topSheet.Name.Contains("Sh") || topSheet.Name.Contains("Ins")) topSheet.Cells["A19"].Value = "Output: " + qtySum + " CUTS, " + Math.Round(areaSum / 1000000, 0) + " m2"; ;
        }

        private static void DetermineTopSheetColorTab(ExcelWorksheet topSheet)
        {
            if(topSheet.Name.Contains("Panel")) topSheet.TabColor = Color.FromArgb(255, 204, 153);
            if(topSheet.Name.Contains("Cut")) topSheet.TabColor = Color.FromArgb(51, 204, 204);
            if (topSheet.Name.Contains("Bat")) topSheet.TabColor = Color.FromArgb(153, 51, 0);
            if (topSheet.Name.Contains("Sh")) topSheet.TabColor = Color.FromArgb(255, 153, 0);
            if (topSheet.Name.Contains("Ins")) topSheet.TabColor = Color.FromArgb(51, 153, 102);
        }

        public static string Treatment(string wsName)
        {
            string treatment;
            if (wsName.Contains("EXT"))
            {
                treatment = "FULL TREATED";
            }
            else
            {
                treatment = "UNTREATED";
            }

            return treatment;
        }

        public static string ItemName(string wsName)
        {
            string itemname = "";
            if (wsName.Contains("Panel")) itemname = "MAKE " + wsName.Remove(0, 3).Replace("Panel", "") + "PANELS";

            if (wsName.Contains("Cut")) itemname = "CUT " + wsName.Remove(0, 3).Replace("Cut", "") + "PANEL CLS";

            if (wsName.Contains("Bat")) itemname = "CUT ULTIMA " + wsName.Remove(0, 3).Replace("Bat", "") + "BATTENS";

            if (wsName.Contains("Sh")) itemname = "CUT " + wsName.Remove(0, 3).Replace("Sh", "") + "PANEL SHEATHING";

            if (wsName.Contains("Ins")) itemname = "CUT " + wsName.Remove(0, 3).Replace("Ins", "") + "PANEL INSULATION";

            return itemname;
        }

    }
}
