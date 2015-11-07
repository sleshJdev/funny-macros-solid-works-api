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
            this.alignWithShaftButton.Click += AlignWithShaftButton_Click;
            this.setupCorpusButton.Click += SetupCorpusButton_Click;
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
            solidWorksHelper.DoProcessSelection();
        }

        private void AddMateButton_Click(object sender, EventArgs e)
        {
            solidWorksHelper.AddMate();
        }

        private void SetupCorpusButton_Click(object sender, EventArgs e)
        {
            solidWorksHelper.SetupCorpus();
        }

        private void AlignWithShaftButton_Click(object sender, EventArgs e)
        {
            solidWorksHelper.AlignWithShaft();
        }

        private void AlignWithHorizont_Click(object sender, EventArgs e)
        {
            solidWorksHelper.AlignAllWithHorizont();
        }

        private void MainForm_FormClosing(object sender, EventArgs e)
        {
            solidWorksHelper.Close();
        }
    }
}
