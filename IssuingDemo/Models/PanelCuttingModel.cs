using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IssuingDemo.Models
{
    class PanelCuttingModel
    {
        public string PanelType { get; set; }
        public string PanelRef { get; set; }
        public string PanelSquareAngled { get; set; }
        public string Material { get; set; }
        public string Grade { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public double Length { get; set; }
        public double LeftCut { get; set; }
        public double RightCut { get; set; }
        public int Qty { get; set; }
    }
}
