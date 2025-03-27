namespace PortiaMoxyImport
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
            this.tbScreen = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.btnPortiaFiles = new System.Windows.Forms.Button();
            this.buttonMoxyToAIM = new System.Windows.Forms.Button();
            this.btnPortiaFilesFromImex = new System.Windows.Forms.Button();
            this.btn_FXConnectTrades = new System.Windows.Forms.Button();
            this.btn_Evare = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.btn_FXTRades_AIM_New = new System.Windows.Forms.Button();
            this.buttonMoxyAIM = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // tbScreen
            // 
            this.tbScreen.Location = new System.Drawing.Point(24, 26);
            this.tbScreen.Multiline = true;
            this.tbScreen.Name = "tbScreen";
            this.tbScreen.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.tbScreen.Size = new System.Drawing.Size(519, 262);
            this.tbScreen.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(186, 369);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(61, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "Test Unit 1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Visible = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(21, 349);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(38, 13);
            this.lblStatus.TabIndex = 2;
            this.lblStatus.Text = "Ready";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(23, 309);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(149, 23);
            this.button2.TabIndex = 3;
            this.button2.Text = "Portia Holdings For Moxy";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // btnPortiaFiles
            // 
            this.btnPortiaFiles.Location = new System.Drawing.Point(240, 369);
            this.btnPortiaFiles.Name = "btnPortiaFiles";
            this.btnPortiaFiles.Size = new System.Drawing.Size(50, 23);
            this.btnPortiaFiles.TabIndex = 4;
            this.btnPortiaFiles.Text = "Portia Files";
            this.btnPortiaFiles.UseVisualStyleBackColor = true;
            this.btnPortiaFiles.Visible = false;
            this.btnPortiaFiles.Click += new System.EventHandler(this.btnPortiaFiles_Click);
            // 
            // buttonMoxyToAIM
            // 
            this.buttonMoxyToAIM.Location = new System.Drawing.Point(190, 397);
            this.buttonMoxyToAIM.Name = "buttonMoxyToAIM";
            this.buttonMoxyToAIM.Size = new System.Drawing.Size(195, 23);
            this.buttonMoxyToAIM.TabIndex = 5;
            this.buttonMoxyToAIM.Text = "MOXY -> AIM Trades";
            this.buttonMoxyToAIM.UseVisualStyleBackColor = true;
            this.buttonMoxyToAIM.Visible = false;
            this.buttonMoxyToAIM.Click += new System.EventHandler(this.buttonMoxyToAIM_Click);
            // 
            // btnPortiaFilesFromImex
            // 
            this.btnPortiaFilesFromImex.Location = new System.Drawing.Point(392, 369);
            this.btnPortiaFilesFromImex.Name = "btnPortiaFilesFromImex";
            this.btnPortiaFilesFromImex.Size = new System.Drawing.Size(140, 23);
            this.btnPortiaFilesFromImex.TabIndex = 6;
            this.btnPortiaFilesFromImex.Text = "Moxy Trades For AIM Old";
            this.btnPortiaFilesFromImex.UseVisualStyleBackColor = true;
            this.btnPortiaFilesFromImex.Visible = false;
            this.btnPortiaFilesFromImex.Click += new System.EventHandler(this.btnPortiaFilesFromImex_Click);
            // 
            // btn_FXConnectTrades
            // 
            this.btn_FXConnectTrades.Location = new System.Drawing.Point(398, 398);
            this.btn_FXConnectTrades.Name = "btn_FXConnectTrades";
            this.btn_FXConnectTrades.Size = new System.Drawing.Size(135, 23);
            this.btn_FXConnectTrades.TabIndex = 7;
            this.btn_FXConnectTrades.Text = "FX Trades -> AIM";
            this.btn_FXConnectTrades.UseVisualStyleBackColor = true;
            this.btn_FXConnectTrades.Visible = false;
            this.btn_FXConnectTrades.Click += new System.EventHandler(this.btn_FXConnectTrades_Click);
            // 
            // btn_Evare
            // 
            this.btn_Evare.Location = new System.Drawing.Point(294, 369);
            this.btn_Evare.Margin = new System.Windows.Forms.Padding(2);
            this.btn_Evare.Name = "btn_Evare";
            this.btn_Evare.Size = new System.Drawing.Size(92, 23);
            this.btn_Evare.TabIndex = 9;
            this.btn_Evare.Text = "Evare --> AIM";
            this.btn_Evare.UseVisualStyleBackColor = true;
            this.btn_Evare.Visible = false;
            this.btn_Evare.Click += new System.EventHandler(this.btn_Evare_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(186, 310);
            this.button3.Margin = new System.Windows.Forms.Padding(2);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(195, 24);
            this.button3.TabIndex = 11;
            this.button3.Text = "Moxy -> AIM Trades";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click_1);
            // 
            // btn_FXTRades_AIM_New
            // 
            this.btn_FXTRades_AIM_New.Location = new System.Drawing.Point(392, 309);
            this.btn_FXTRades_AIM_New.Margin = new System.Windows.Forms.Padding(2);
            this.btn_FXTRades_AIM_New.Name = "btn_FXTRades_AIM_New";
            this.btn_FXTRades_AIM_New.Size = new System.Drawing.Size(135, 23);
            this.btn_FXTRades_AIM_New.TabIndex = 13;
            this.btn_FXTRades_AIM_New.Text = "FXTrades -> AIM ";
            this.btn_FXTRades_AIM_New.UseVisualStyleBackColor = true;
            this.btn_FXTRades_AIM_New.Click += new System.EventHandler(this.btn_FXTRades_AIM_New_Click);
            // 
            // buttonMoxyAIM
            // 
            this.buttonMoxyAIM.Location = new System.Drawing.Point(228, 339);
            this.buttonMoxyAIM.Name = "buttonMoxyAIM";
            this.buttonMoxyAIM.Size = new System.Drawing.Size(75, 23);
            this.buttonMoxyAIM.TabIndex = 14;
            this.buttonMoxyAIM.Text = "Moxy AIM Test";
            this.buttonMoxyAIM.UseVisualStyleBackColor = true;
            this.buttonMoxyAIM.Visible = false;
            this.buttonMoxyAIM.Click += new System.EventHandler(this.buttonMoxyAIM_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(562, 444);
            this.Controls.Add(this.buttonMoxyAIM);
            this.Controls.Add(this.btn_FXTRades_AIM_New);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.btn_Evare);
            this.Controls.Add(this.btn_FXConnectTrades);
            this.Controls.Add(this.btnPortiaFilesFromImex);
            this.Controls.Add(this.buttonMoxyToAIM);
            this.Controls.Add(this.btnPortiaFiles);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.tbScreen);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Portia Moxy Import";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tbScreen;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button btnPortiaFiles;
        private System.Windows.Forms.Button buttonMoxyToAIM;
        private System.Windows.Forms.Button btnPortiaFilesFromImex;
        private System.Windows.Forms.Button btn_FXConnectTrades;
        private System.Windows.Forms.Button btn_Evare;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button btn_FXTRades_AIM_New;
        private System.Windows.Forms.Button buttonMoxyAIM;
    }
}

