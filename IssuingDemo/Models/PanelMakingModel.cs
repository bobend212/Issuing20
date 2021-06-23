using CsvHelper.Configuration;
using OfficeOpenXml.Attributes;
using System.Collections;

namespace IssuingDemo
{
    public class PanelMakingModel
    {
        public string PanelType { get; set; }
        public string PanelRef { get; set; }
        public string PanelSquareAngled { get; set; }
        public double Length { get; set; }
        public double Height { get; set; }
        public double Area { get; set; }
        public double Weight { get; set; }
        public int Qty { get; set; }
    }

}
