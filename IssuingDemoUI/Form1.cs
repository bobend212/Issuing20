using IssuingDemo;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IssuingDemoUI
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }



        private async void button1_Click_1(object sender, EventArgs e)
        {
            PanelMaking panelMaking = new PanelMaking();
            PanelCutting panelCutting = new PanelCutting();

            panelMaking._clientName = txtClientName.Text;
            panelMaking._jobNo = txtJobNo.Text;
            panelMaking._plot = txtPlot.Text;
            panelMaking._setRef = cbSetRef.Text;

            panelCutting._clientName = txtClientName.Text;
            panelCutting._jobNo = txtJobNo.Text;
            panelCutting._plot = txtPlot.Text;
            panelCutting._setRef = cbSetRef.Text;

            //cutting
            if (cbCutting.Checked) await panelCutting.GeneratePanelCutting();
            //making
            if (cbPanels.Checked) await panelMaking.GeneratePanelMaking();


        }



    }
}
