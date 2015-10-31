using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace FunnyMacros
{
    public partial class MainForm : Form
    {
        private SolidWorksHelper swh = new SolidWorksHelper();

        



        public MainForm()
        {            
            InitializeComponent();
            swh.Close();
            this.runSwButton.Click += RunSolidWorksButton_Click;
            this.processSelectionButton.Click += ProcessSelectionButton_Click;
            this.loadLocatorsButton.Click += LoadLocatorsButton_Click;
            this.addMateButton.Click += AddMateButton_Click;
            this.loadShaftButton.Click += LoadShaftButton_Click;
            this.FormClosing += MainForm_FormClosing;
        }
        
        private void RunSolidWorksButton_Click(object sender, EventArgs e)
        {
            swh.Open();
            swh.Initalize();
        }

        private void ProcessSelectionButton_Click(object sender, EventArgs e)
        {
            try
            {
                swh.DoProcessSelection();                
            }
            catch (Exception ex)
            {
                MessageBox.Show(null, string.Format("Some problems occured during execute. \nDetails: {0}", ex.Message), "Problem",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3, MessageBoxOptions.DefaultDesktopOnly);
            }

        }       

        private void AddMateButton_Click(object sender, EventArgs e)
        {
            swh.AddMate();
        }

        private void LoadLocatorsButton_Click(object sender, EventArgs e)
        {
            


        }

        private void LoadShaftButton_Click(object sender, EventArgs e)
        {             
        }

        private void MainForm_FormClosing(object sender, EventArgs e)
        {
            swh.Close();
        }         
    }
}
