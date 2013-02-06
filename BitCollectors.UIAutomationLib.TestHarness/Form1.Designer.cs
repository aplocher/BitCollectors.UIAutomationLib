namespace BitCollectors.UIAutomationLib.TestHarness
{
    partial class lblAutomationFile
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
            this.btnRunTest = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btnOpenAutomationFile = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnRunTest
            // 
            this.btnRunTest.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRunTest.Location = new System.Drawing.Point(141, 147);
            this.btnRunTest.Name = "btnRunTest";
            this.btnRunTest.Size = new System.Drawing.Size(131, 23);
            this.btnRunTest.TabIndex = 4;
            this.btnRunTest.Text = "Execute Automation";
            this.btnRunTest.UseVisualStyleBackColor = true;
            this.btnRunTest.Click += new System.EventHandler(this.btnRunTest_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnOpenAutomationFile
            // 
            this.btnOpenAutomationFile.Location = new System.Drawing.Point(17, 38);
            this.btnOpenAutomationFile.Name = "btnOpenAutomationFile";
            this.btnOpenAutomationFile.Size = new System.Drawing.Size(75, 23);
            this.btnOpenAutomationFile.TabIndex = 1;
            this.btnOpenAutomationFile.Text = "Browse";
            this.btnOpenAutomationFile.UseVisualStyleBackColor = true;
            this.btnOpenAutomationFile.Click += new System.EventHandler(this.btnOpenAutomationFile_Click);
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(98, 43);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(182, 18);
            this.label1.TabIndex = 0;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(17, 99);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(255, 20);
            this.textBox1.TabIndex = 3;
            this.textBox1.Text = "hello world - abc 123";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(14, 78);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(69, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Sample Text:";
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(14, 17);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(182, 18);
            this.label3.TabIndex = 5;
            this.label3.Text = "Select the automation config file";
            // 
            // lblAutomationFile
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 182);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnOpenAutomationFile);
            this.Controls.Add(this.btnRunTest);
            this.Name = "lblAutomationFile";
            this.Text = "GenericAutomation Test Application";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnRunTest;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btnOpenAutomationFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
    }
}

