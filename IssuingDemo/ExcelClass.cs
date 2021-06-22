using CsvHelper;
using CsvHelper.Configuration;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace IssuingDemo
{
     public static class ExcelClass
    {



        private static List<PanelModel> GetSetupData()
        {
            List<PanelModel> output = new()
            {
                new() { PanelType = "Ext", PanelRef = "GF-XD1", PanelSquareAngled = "Sq", Height = 2300, Length = 2200, Weight = 45, Area = 2.3, Qty = 1 },
                new() { PanelType = "Ext", PanelRef = "GF-XD2", PanelSquareAngled = "Sq", Height = 2300, Length = 220, Weight = 43, Area = 2.3, Qty = 1 },
                new() { PanelType = "Ext", PanelRef = "GF-XD3", PanelSquareAngled = "Sq", Height = 2100, Length = 2250, Weight = 41, Area = 6.3, Qty = 1 },
                new() { PanelType = "Int", PanelRef = "GF-ND1", PanelSquareAngled = "Sq", Height = 2345, Length = 1000, Weight = 40, Area = 1.3, Qty = 1 },
                new() { PanelType = "Int", PanelRef = "GF-ND2", PanelSquareAngled = "Sq", Height = 1500, Length = 567, Weight = 33, Area = 1.1, Qty = 1 }
            };

            return output;
        }
    }
}
