using System;
using System.Windows.Forms;
using FunnyMacros.Util;

namespace FunnyMacros
{
    public partial class MainForm : Form
    {

        private SolidWorksMacro solidWorksHelper = new SolidWorksMacro();

        public MainForm()
        {            
            InitializeComponent();
            solidWorksHelper.Close();
            this.runSwButton.Click += RunSolidWorksButton_Click;
            this.processSelectionButton.Click += ProcessSelectionButton_Click;
            this.alignWithHorizontal.Click += AlignWithHorizont_Click;
            this.addMateButton.Click += AddMateButton_Click;
            this.FormClosing += MainForm_FormClosing;
        }
        
        private void RunSolidWorksButton_Click(object sender, EventArgs e)
        {
            solidWorksHelper.Open();
            solidWorksHelper.Initalize();
        }

        private void ProcessSelectionButton_Click(object sender, EventArgs e)
        {
            try
            {
                solidWorksHelper.DoProcessSelection();                
            }
            catch (Exception ex)
            {
                MessageBox.Show(null, string.Format("Some problems occured during execute. \nDetails: {0}", ex.Message), "Problem",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3, MessageBoxOptions.DefaultDesktopOnly);
            }

        }       

        private void AddMateButton_Click(object sender, EventArgs e)
        {
            solidWorksHelper.AddMate();
        }

        private void AlignWithHorizont_Click(object sender, EventArgs e)
        {
            solidWorksHelper.AlignWithHorizontAll();
        }

        private void MainForm_FormClosing(object sender, EventArgs e)
        {
            solidWorksHelper.Close();
        }         
    }
}
