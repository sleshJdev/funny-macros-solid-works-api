using System;
using System.Windows.Forms;
using FunnyMacros.Macros;

namespace FunnyMacros
{
    public partial class MainForm : Form
    {
        private SolidWorksMacro solidWorksHelper = new SolidWorksMacro();

        public MainForm()
        {
            InitializeComponent();
            runSwButton.Click               += RunSolidWorksButton_Click;
            checkSizeAccordanceButton.Click += CheckSizeAccordanceButton_Click;
            processSelectionButton.Click    += ProcessSelectionButton_Click;
            alignWithHorizontal.Click       += AlignWithHorizont_Click;
            alignWithShaftButton.Click      += AlignWithShaftButton_Click;
            setupCorpusButton.Click         += SetupCorpusButton_Click;
            addMateButton.Click             += AddMateButton_Click;
            adjustSizesButton.Click         += AdjustSizesButton_Click;
            FormClosing                     += MainForm_FormClosing;
        }

        private void CheckSizeAccordanceButton_Click(object sender, EventArgs e)
        {
            solidWorksHelper.CheckSizeAccordance();
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

        private void AdjustSizesButton_Click(object sender, EventArgs e)
        {
            solidWorksHelper.AdjustSize();
        }

        private void MainForm_FormClosing(object sender, EventArgs e)
        {
            //solidWorksHelper.Close();
        }
    }
}
