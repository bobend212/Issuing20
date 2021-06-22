using CsvHelper.Configuration;
using OfficeOpenXml.Attributes;

namespace IssuingDemo
{
    public class PanelModel
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

    public class PanelModelMapped : ClassMap<PanelModel>
    {
        public string PanelRef { get; set; }
        public double Length { get; set; }
        public double Height { get; set; }
        public double Area { get; set; }
        public double Weight { get; set; }
        public int Qty { get; set; }
    }

}
