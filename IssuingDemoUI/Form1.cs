using IssuingDemo;
using IssuingDemoLogger;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
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

        private async void btnRunIssuing_Click_1(object sender, EventArgs e)
        {
            PanelMaking panelMaking = new PanelMaking();
            PanelCutting panelCutting = new PanelCutting();
            PanelSheathing panelSheathing = new PanelSheathing();
            PanelInsulation panelInsulation = new PanelInsulation();

            panelMaking._clientName = txtClientName.Text;
            panelMaking._jobNo = txtJobNo.Text;
            panelMaking._plot = txtPlot.Text;
            panelMaking._setRef = cbSetRef.Text;
            panelMaking._mbaJob = txtMbaJob.Text.Trim();

            panelCutting._clientName = txtClientName.Text;
            panelCutting._jobNo = txtJobNo.Text;
            panelCutting._plot = txtPlot.Text;
            panelCutting._setRef = cbSetRef.Text;
            panelCutting._mbaJob = txtMbaJob.Text.Trim();

            panelSheathing._clientName = txtClientName.Text;
            panelSheathing._jobNo = txtJobNo.Text;
            panelSheathing._plot = txtPlot.Text;
            panelSheathing._setRef = cbSetRef.Text;
            panelSheathing._mbaJob = txtMbaJob.Text.Trim();

            panelInsulation._clientName = txtClientName.Text;
            panelInsulation._jobNo = txtJobNo.Text;
            panelInsulation._plot = txtPlot.Text;
            panelInsulation._setRef = cbSetRef.Text;
            panelInsulation._mbaJob = txtMbaJob.Text.Trim();

            try
            {
                //insulation
                if (cbInsulation.Checked) await panelInsulation.GeneratePanelInsulation();
                //sheathing
                if (cbSheathing.Checked) await panelSheathing.GeneratePanelSheathing();
                //cutting
                if (cbCutting.Checked) await panelCutting.GeneratePanelCutting();
                //making
                if (cbPanels.Checked) await panelMaking.GeneratePanelMaking();
            }       
            catch (Exception ex)
            {
                var logger = new FileLoggerBuilder().WithPath("logs.txt").WithDate().Build();
                logger.Log(ex.Message);
                MessageBox.Show(ex.Message);
            }

        }

        private void btnOpenIssuing_Click(object sender, EventArgs e)
        {
            string pathToIssuing = $@"C:\MiTek\UK\jobs\{txtMbaJob.Text.Trim()}\Attachments\{txtJobNo.Text}_issuing.xlsx";
            Process.Start(new ProcessStartInfo(pathToIssuing) { UseShellExecute = true });
        }
    }
}
