
namespace IssuingDemoUI
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.txtClientName = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnOpenIssuing = new System.Windows.Forms.Button();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.label2 = new System.Windows.Forms.Label();
            this.txtJobNo = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtPlot = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cbSetRef = new System.Windows.Forms.ComboBox();
            this.btnRunIssuing = new System.Windows.Forms.Button();
            this.cbPanels = new System.Windows.Forms.CheckBox();
            this.cbCutting = new System.Windows.Forms.CheckBox();
            this.cbSheathing = new System.Windows.Forms.CheckBox();
            this.cbInsulation = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.flowLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(73, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "Client Name";
            // 
            // txtClientName
            // 
            this.txtClientName.Location = new System.Drawing.Point(3, 18);
            this.txtClientName.Name = "txtClientName";
            this.txtClientName.Size = new System.Drawing.Size(250, 23);
            this.txtClientName.TabIndex = 1;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnOpenIssuing);
            this.groupBox1.Controls.Add(this.flowLayoutPanel1);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(285, 426);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Data";
            // 
            // btnOpenIssuing
            // 
            this.btnOpenIssuing.BackColor = System.Drawing.Color.DeepSkyBlue;
            this.btnOpenIssuing.Location = new System.Drawing.Point(9, 348);
            this.btnOpenIssuing.Name = "btnOpenIssuing";
            this.btnOpenIssuing.Size = new System.Drawing.Size(113, 51);
            this.btnOpenIssuing.TabIndex = 8;
            this.btnOpenIssuing.Text = "Open Issuing";
            this.btnOpenIssuing.UseVisualStyleBackColor = false;
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Controls.Add(this.label1);
            this.flowLayoutPanel1.Controls.Add(this.txtClientName);
            this.flowLayoutPanel1.Controls.Add(this.label2);
            this.flowLayoutPanel1.Controls.Add(this.txtJobNo);
            this.flowLayoutPanel1.Controls.Add(this.label3);
            this.flowLayoutPanel1.Controls.Add(this.txtPlot);
            this.flowLayoutPanel1.Controls.Add(this.label4);
            this.flowLayoutPanel1.Controls.Add(this.cbSetRef);
            this.flowLayoutPanel1.Location = new System.Drawing.Point(6, 22);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(263, 193);
            this.flowLayoutPanel1.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 44);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 15);
            this.label2.TabIndex = 2;
            this.label2.Text = "Job No.";
            // 
            // txtJobNo
            // 
            this.txtJobNo.Location = new System.Drawing.Point(3, 62);
            this.txtJobNo.Name = "txtJobNo";
            this.txtJobNo.Size = new System.Drawing.Size(250, 23);
            this.txtJobNo.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(3, 88);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(28, 15);
            this.label3.TabIndex = 4;
            this.label3.Text = "Plot";
            // 
            // txtPlot
            // 
            this.txtPlot.Location = new System.Drawing.Point(3, 106);
            this.txtPlot.Name = "txtPlot";
            this.txtPlot.Size = new System.Drawing.Size(250, 23);
            this.txtPlot.TabIndex = 5;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(3, 132);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(46, 15);
            this.label4.TabIndex = 6;
            this.label4.Text = "Set Ref.";
            // 
            // cbSetRef
            // 
            this.cbSetRef.FormattingEnabled = true;
            this.cbSetRef.Items.AddRange(new object[] {
            "GF Walls",
            "FF Walls",
            "SF Walls",
            "FF Joists",
            "SF Joists",
            "Roof Joists",
            "Roof"});
            this.cbSetRef.Location = new System.Drawing.Point(3, 150);
            this.cbSetRef.Name = "cbSetRef";
            this.cbSetRef.Size = new System.Drawing.Size(250, 23);
            this.cbSetRef.TabIndex = 7;
            // 
            // btnRunIssuing
            // 
            this.btnRunIssuing.BackColor = System.Drawing.Color.MediumSpringGreen;
            this.btnRunIssuing.Location = new System.Drawing.Point(478, 360);
            this.btnRunIssuing.Name = "btnRunIssuing";
            this.btnRunIssuing.Size = new System.Drawing.Size(113, 51);
            this.btnRunIssuing.TabIndex = 4;
            this.btnRunIssuing.Text = "Run Issuing";
            this.btnRunIssuing.UseVisualStyleBackColor = false;
            this.btnRunIssuing.Click += new System.EventHandler(this.btnRunIssuing_Click_1);
            // 
            // cbPanels
            // 
            this.cbPanels.AutoSize = true;
            this.cbPanels.Location = new System.Drawing.Point(315, 52);
            this.cbPanels.Name = "cbPanels";
            this.cbPanels.Size = new System.Drawing.Size(60, 19);
            this.cbPanels.TabIndex = 5;
            this.cbPanels.Text = "Panels";
            this.cbPanels.UseVisualStyleBackColor = true;
            // 
            // cbCutting
            // 
            this.cbCutting.AutoSize = true;
            this.cbCutting.Location = new System.Drawing.Point(315, 78);
            this.cbCutting.Name = "cbCutting";
            this.cbCutting.Size = new System.Drawing.Size(66, 19);
            this.cbCutting.TabIndex = 6;
            this.cbCutting.Text = "Cutting";
            this.cbCutting.UseVisualStyleBackColor = true;
            // 
            // cbSheathing
            // 
            this.cbSheathing.AutoSize = true;
            this.cbSheathing.Location = new System.Drawing.Point(315, 103);
            this.cbSheathing.Name = "cbSheathing";
            this.cbSheathing.Size = new System.Drawing.Size(79, 19);
            this.cbSheathing.TabIndex = 7;
            this.cbSheathing.Text = "Sheathing";
            this.cbSheathing.UseVisualStyleBackColor = true;
            // 
            // cbInsulation
            // 
            this.cbInsulation.AutoSize = true;
            this.cbInsulation.Location = new System.Drawing.Point(315, 128);
            this.cbInsulation.Name = "cbInsulation";
            this.cbInsulation.Size = new System.Drawing.Size(78, 19);
            this.cbInsulation.TabIndex = 8;
            this.cbInsulation.Text = "Insulation";
            this.cbInsulation.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(603, 423);
            this.Controls.Add(this.cbInsulation);
            this.Controls.Add(this.cbSheathing);
            this.Controls.Add(this.cbCutting);
            this.Controls.Add(this.cbPanels);
            this.Controls.Add(this.btnRunIssuing);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Issuing 2.0";
            this.groupBox1.ResumeLayout(false);
            this.flowLayoutPanel1.ResumeLayout(false);
            this.flowLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtClientName;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtJobNo;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtPlot;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cbSetRef;
        private System.Windows.Forms.Button btnRunIssuing;
        private System.Windows.Forms.CheckBox cbPanels;
        private System.Windows.Forms.CheckBox cbCutting;
        private System.Windows.Forms.CheckBox cbSheathing;
        private System.Windows.Forms.Button btnOpenIssuing;
        private System.Windows.Forms.CheckBox cbInsulation;
    }
}