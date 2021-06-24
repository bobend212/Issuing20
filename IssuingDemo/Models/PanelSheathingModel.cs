using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IssuingDemo.Models
{
    class PanelSheathingModel
    {
        public string PanelType { get; set; }
        public string PanelRef { get; set; }
        public string PanelSquareAngled { get; set; }
        public string Type { get; set; }
        public string Material { get; set; }
        public double Thickness { get; set; }
        public double Height { get; set; }
        public double Width { get; set; }
        public int Qty { get; set; }
    }
}
