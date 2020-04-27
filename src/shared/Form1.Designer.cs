namespace PanelScheduleExporter
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
      this.components = new System.ComponentModel.Container();
      System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
      this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
      this.button1 = new System.Windows.Forms.Button();
      this.textBox1 = new System.Windows.Forms.TextBox();
      this.label1 = new System.Windows.Forms.Label();
      this.btnDirectoryPick = new System.Windows.Forms.Button();
      this.progressBar1 = new System.Windows.Forms.ProgressBar();
      this.btnCheckAll = new System.Windows.Forms.Button();
      this.btnCheckNone = new System.Windows.Forms.Button();
      this.label2 = new System.Windows.Forms.Label();
      this.timer1 = new System.Windows.Forms.Timer(this.components);
      this.lblProgress = new System.Windows.Forms.Label();
      this.btnCancel = new System.Windows.Forms.Button();
      this.SuspendLayout();
      // 
      // checkedListBox1
      // 
      this.checkedListBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
      this.checkedListBox1.CheckOnClick = true;
      this.checkedListBox1.FormattingEnabled = true;
      this.checkedListBox1.Location = new System.Drawing.Point(12, 27);
      this.checkedListBox1.Name = "checkedListBox1";
      this.checkedListBox1.Size = new System.Drawing.Size(145, 154);
      this.checkedListBox1.TabIndex = 0;
      // 
      // button1
      // 
      this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
      this.button1.Location = new System.Drawing.Point(322, 169);
      this.button1.Name = "button1";
      this.button1.Size = new System.Drawing.Size(99, 42);
      this.button1.TabIndex = 1;
      this.button1.Text = "Export";
      this.button1.UseVisualStyleBackColor = true;
      this.button1.Click += new System.EventHandler(this.button1_Click);
      // 
      // textBox1
      // 
      this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
      this.textBox1.Enabled = false;
      this.textBox1.Location = new System.Drawing.Point(179, 29);
      this.textBox1.Name = "textBox1";
      this.textBox1.Size = new System.Drawing.Size(210, 20);
      this.textBox1.TabIndex = 2;
      // 
      // label1
      // 
      this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
      this.label1.AutoSize = true;
      this.label1.Location = new System.Drawing.Point(179, 9);
      this.label1.Name = "label1";
      this.label1.Size = new System.Drawing.Size(145, 13);
      this.label1.TabIndex = 3;
      this.label1.Text = "Choose directory to export to:";
      // 
      // btnDirectoryPick
      // 
      this.btnDirectoryPick.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
      this.btnDirectoryPick.Location = new System.Drawing.Point(395, 26);
      this.btnDirectoryPick.Name = "btnDirectoryPick";
      this.btnDirectoryPick.Size = new System.Drawing.Size(26, 23);
      this.btnDirectoryPick.TabIndex = 4;
      this.btnDirectoryPick.Text = "...";
      this.btnDirectoryPick.UseVisualStyleBackColor = true;
      this.btnDirectoryPick.Click += new System.EventHandler(this.btnDirectoryPick_Click);
      // 
      // progressBar1
      // 
      this.progressBar1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
      this.progressBar1.Location = new System.Drawing.Point(179, 96);
      this.progressBar1.Name = "progressBar1";
      this.progressBar1.Size = new System.Drawing.Size(242, 23);
      this.progressBar1.TabIndex = 5;
      // 
      // btnCheckAll
      // 
      this.btnCheckAll.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
      this.btnCheckAll.Location = new System.Drawing.Point(12, 188);
      this.btnCheckAll.Name = "btnCheckAll";
      this.btnCheckAll.Size = new System.Drawing.Size(61, 23);
      this.btnCheckAll.TabIndex = 6;
      this.btnCheckAll.Text = "Check All";
      this.btnCheckAll.UseVisualStyleBackColor = true;
      this.btnCheckAll.Click += new System.EventHandler(this.btnCheckAll_Click);
      // 
      // btnCheckNone
      // 
      this.btnCheckNone.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
      this.btnCheckNone.Location = new System.Drawing.Point(79, 188);
      this.btnCheckNone.Name = "btnCheckNone";
      this.btnCheckNone.Size = new System.Drawing.Size(78, 23);
      this.btnCheckNone.TabIndex = 6;
      this.btnCheckNone.Text = "Check None";
      this.btnCheckNone.UseVisualStyleBackColor = true;
      this.btnCheckNone.Click += new System.EventHandler(this.btnCheckNone_Click);
      // 
      // label2
      // 
      this.label2.AutoSize = true;
      this.label2.Location = new System.Drawing.Point(12, 9);
      this.label2.Name = "label2";
      this.label2.Size = new System.Drawing.Size(134, 13);
      this.label2.TabIndex = 3;
      this.label2.Text = "Panel Schedules in Project";
      // 
      // timer1
      // 
      this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
      // 
      // lblProgress
      // 
      this.lblProgress.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.lblProgress.AutoSize = true;
      this.lblProgress.Location = new System.Drawing.Point(179, 78);
      this.lblProgress.Name = "lblProgress";
      this.lblProgress.Size = new System.Drawing.Size(13, 13);
      this.lblProgress.TabIndex = 7;
      this.lblProgress.Text = "--";
      this.lblProgress.Visible = false;
      // 
      // btnCancel
      // 
      this.btnCancel.Location = new System.Drawing.Point(241, 188);
      this.btnCancel.Name = "btnCancel";
      this.btnCancel.Size = new System.Drawing.Size(75, 23);
      this.btnCancel.TabIndex = 8;
      this.btnCancel.Text = "Cancel";
      this.btnCancel.UseVisualStyleBackColor = true;
      this.btnCancel.Visible = false;
      this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
      // 
      // Form1
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(433, 226);
      this.Controls.Add(this.btnCancel);
      this.Controls.Add(this.lblProgress);
      this.Controls.Add(this.btnCheckNone);
      this.Controls.Add(this.btnCheckAll);
      this.Controls.Add(this.progressBar1);
      this.Controls.Add(this.btnDirectoryPick);
      this.Controls.Add(this.label2);
      this.Controls.Add(this.label1);
      this.Controls.Add(this.textBox1);
      this.Controls.Add(this.button1);
      this.Controls.Add(this.checkedListBox1);
      this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
      this.MinimumSize = new System.Drawing.Size(449, 264);
      this.Name = "Form1";
      this.Text = "Export Panel Schedules to Excel";
      this.Load += new System.EventHandler(this.Form1_Load);
      this.ResumeLayout(false);
      this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnDirectoryPick;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button btnCheckAll;
        private System.Windows.Forms.Button btnCheckNone;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Label lblProgress;
        private System.Windows.Forms.Button btnCancel;
    }
}