using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;

namespace PanelScheduleExporter
{
    /// <summary>
    /// Form
    /// </summary>
    public partial class Form1 : System.Windows.Forms.Form
    {
        /// <summary>
        /// 
        /// </summary>
        public Document _doc;

        /// <summary>
        /// 
        /// </summary>
        private int intProgress = 0;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="doc"></param>
        public Form1(Document doc)
        {
            InitializeComponent();
            _doc = doc;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            progressBar1.Visible = false;

            textBox1.Text = PanelScheduleExport._exportDirectory;            

            FilteredElementCollector fec = new FilteredElementCollector(_doc).OfClass(typeof(PanelScheduleView));
            
            foreach (Element elem in fec)
            {
                PanelScheduleView psView = elem as PanelScheduleView;
                if (psView.IsPanelScheduleTemplate())
                {
                    continue;
                }
                else
                {
                    checkedListBox1.Items.Add(psView, false);
                }
            }
            checkedListBox1.DisplayMember = "Name";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                //Set buttons disabled
                button1.Enabled = false;
                btnCheckAll.Enabled = false;
                btnCheckNone.Enabled = false;
                btnDirectoryPick.Enabled = false;
                btnCancel.Visible = true;
                lblProgress.Visible = true;

                //Progress bar
                progressBar1.Visible = true;
                progressBar1.Enabled = true;
                progressBar1.Style = ProgressBarStyle.Continuous;
                progressBar1.Maximum = 100;
                progressBar1.Value = 0;
                timer1.Enabled = true;                               

                backgroundWorker1.RunWorkerAsync();
                
            }
            catch (IOException)
            {
                MessageBox.Show("Export operation was canceled before completing.");
                this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            }
            catch (Exception)
            {
                string message = "You probably do not have Microsoft Excel installed on this system.\nExporter cannot work without this.";
                MessageBox.Show(message);
                this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            }
        }

        private void btnDirectoryPick_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();            
            fbd.ShowNewFolderButton = true;
            DialogResult ddr = fbd.ShowDialog();

            if (ddr == System.Windows.Forms.DialogResult.OK)
            {
                PanelScheduleExport._exportDirectory = fbd.SelectedPath;
                textBox1.Text = PanelScheduleExport._exportDirectory;
            }
        }

        private void btnCheckAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, true);
            }
        }

        private void btnCheckNone_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, false);
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (progressBar1.Value < progressBar1.Maximum)
            {
                progressBar1.Increment(5);
            }
            else
            {
                progressBar1.Value = progressBar1.Minimum;
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            lblProgress.Visible = true; 
            foreach (object o in checkedListBox1.CheckedItems)
            {
                lblProgress.Text = string.Format("{0} of {1} Panels processed", intProgress, checkedListBox1.CheckedItems.Count.ToString());
                //PanelScheduleExport._schedulesToExport.Add(o as Element);
                Translator trans = new XLSXTranslator(o as PanelScheduleView) as Translator;
                string exported = trans.Export();

                backgroundWorker1.ReportProgress((int)(intProgress/checkedListBox1.CheckedItems.Count));
                
                if (backgroundWorker1.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }
                intProgress++;                
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //progressBar1.Value = e.ProgressPercentage;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            btnCancel.Enabled = false;
            backgroundWorker1.CancelAsync();
        }
    }
}
