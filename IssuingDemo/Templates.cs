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

namespace IssuingDemo
{
    public class Templates
    {
        public string _clientName;
        public string _jobNo;
        public string _plot;
        public string _setRef;

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
    }
}
