namespace Y2Y_Comparison
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
            this.Browse1 = new System.Windows.Forms.Button();
            this.Browse2 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.StartButton = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.CurrentText = new System.Windows.Forms.Label();
            this.PrevText = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // Browse1
            // 
            this.Browse1.Location = new System.Drawing.Point(250, 21);
            this.Browse1.Name = "Browse1";
            this.Browse1.Size = new System.Drawing.Size(75, 23);
            this.Browse1.TabIndex = 0;
            this.Browse1.Text = "Browse";
            this.Browse1.UseVisualStyleBackColor = true;
            this.Browse1.Click += new System.EventHandler(this.Browse1_Click);
            // 
            // Browse2
            // 
            this.Browse2.Location = new System.Drawing.Point(250, 75);
            this.Browse2.Name = "Browse2";
            this.Browse2.Size = new System.Drawing.Size(75, 23);
            this.Browse2.TabIndex = 1;
            this.Browse2.Text = "Browse";
            this.Browse2.UseVisualStyleBackColor = true;
            this.Browse2.Click += new System.EventHandler(this.Browse2_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label1.Location = new System.Drawing.Point(28, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(208, 17);
            this.label1.TabIndex = 2;
            this.label1.Text = "Choose Current FWG Database";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label2.Location = new System.Drawing.Point(28, 80);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(216, 17);
            this.label2.TabIndex = 3;
            this.label2.Text = "Choose Previous FWG Database";
            // 
            // StartButton
            // 
            this.StartButton.Location = new System.Drawing.Point(130, 130);
            this.StartButton.Name = "StartButton";
            this.StartButton.Size = new System.Drawing.Size(75, 23);
            this.StartButton.TabIndex = 4;
            this.StartButton.Text = "Start";
            this.StartButton.UseVisualStyleBackColor = true;
            this.StartButton.Click += new System.EventHandler(this.StartButton_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.FileName = "openFileDialog2";
            // 
            // CurrentText
            // 
            this.CurrentText.AutoSize = true;
            this.CurrentText.Location = new System.Drawing.Point(49, 43);
            this.CurrentText.Name = "CurrentText";
            this.CurrentText.Size = new System.Drawing.Size(35, 13);
            this.CurrentText.TabIndex = 5;
            this.CurrentText.Text = "label3";
            // 
            // PrevText
            // 
            this.PrevText.AutoSize = true;
            this.PrevText.Location = new System.Drawing.Point(52, 101);
            this.PrevText.Name = "PrevText";
            this.PrevText.Size = new System.Drawing.Size(35, 13);
            this.PrevText.TabIndex = 6;
            this.PrevText.Text = "label3";
            // 
            // folderBrowserDialog1
            // 
            this.folderBrowserDialog1.Description = "Choose Folder to Save Excel";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(352, 176);
            this.Controls.Add(this.PrevText);
            this.Controls.Add(this.CurrentText);
            this.Controls.Add(this.StartButton);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Browse2);
            this.Controls.Add(this.Browse1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Browse1;
        private System.Windows.Forms.Button Browse2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button StartButton;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
        private System.Windows.Forms.Label CurrentText;
        private System.Windows.Forms.Label PrevText;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
    }
}

