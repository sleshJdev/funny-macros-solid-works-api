namespace FunnyMacros
{
    partial class MainForm
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
            this.runSwButton = new System.Windows.Forms.Button();
            this.addMateButton = new System.Windows.Forms.Button();
            this.processSelectionButton = new System.Windows.Forms.Button();
            this.alignWithHorizontal = new System.Windows.Forms.Button();
            this.alignWithShaftButton = new System.Windows.Forms.Button();
            this.setupCorpusButton = new System.Windows.Forms.Button();
            this.adjustSizesButton = new System.Windows.Forms.Button();
            this.checkSizeAccordanceButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // runSwButton
            // 
            this.runSwButton.Location = new System.Drawing.Point(12, 12);
            this.runSwButton.Name = "runSwButton";
            this.runSwButton.Size = new System.Drawing.Size(187, 43);
            this.runSwButton.TabIndex = 0;
            this.runSwButton.Text = "Run SolidWorks. Load Components.";
            this.runSwButton.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.runSwButton.UseVisualStyleBackColor = true;
            // 
            // addMateButton
            // 
            this.addMateButton.Location = new System.Drawing.Point(12, 120);
            this.addMateButton.Name = "addMateButton";
            this.addMateButton.Size = new System.Drawing.Size(187, 25);
            this.addMateButton.TabIndex = 1;
            this.addMateButton.Text = "Add Mate";
            this.addMateButton.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.addMateButton.UseVisualStyleBackColor = true;
            // 
            // processSelectionButton
            // 
            this.processSelectionButton.Location = new System.Drawing.Point(12, 92);
            this.processSelectionButton.Name = "processSelectionButton";
            this.processSelectionButton.Size = new System.Drawing.Size(187, 25);
            this.processSelectionButton.TabIndex = 3;
            this.processSelectionButton.Text = "Process selection";
            this.processSelectionButton.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.processSelectionButton.UseVisualStyleBackColor = true;
            // 
            // alignWithHorizontal
            // 
            this.alignWithHorizontal.Location = new System.Drawing.Point(12, 151);
            this.alignWithHorizontal.Name = "alignWithHorizontal";
            this.alignWithHorizontal.Size = new System.Drawing.Size(187, 25);
            this.alignWithHorizontal.TabIndex = 4;
            this.alignWithHorizontal.Text = "Align With Horizont";
            this.alignWithHorizontal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.alignWithHorizontal.UseVisualStyleBackColor = true;
            // 
            // alignWithShaftButton
            // 
            this.alignWithShaftButton.Location = new System.Drawing.Point(12, 182);
            this.alignWithShaftButton.Name = "alignWithShaftButton";
            this.alignWithShaftButton.Size = new System.Drawing.Size(187, 25);
            this.alignWithShaftButton.TabIndex = 5;
            this.alignWithShaftButton.Text = "Align With Shaft";
            this.alignWithShaftButton.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.alignWithShaftButton.UseVisualStyleBackColor = true;
            // 
            // setupCorpusButton
            // 
            this.setupCorpusButton.Location = new System.Drawing.Point(12, 213);
            this.setupCorpusButton.Name = "setupCorpusButton";
            this.setupCorpusButton.Size = new System.Drawing.Size(187, 25);
            this.setupCorpusButton.TabIndex = 6;
            this.setupCorpusButton.Text = "Setup Corpus";
            this.setupCorpusButton.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.setupCorpusButton.UseVisualStyleBackColor = true;
            // 
            // adjustSizesButton
            // 
            this.adjustSizesButton.Location = new System.Drawing.Point(12, 244);
            this.adjustSizesButton.Name = "adjustSizesButton";
            this.adjustSizesButton.Size = new System.Drawing.Size(187, 25);
            this.adjustSizesButton.TabIndex = 7;
            this.adjustSizesButton.Text = "Adjust Corpus Size";
            this.adjustSizesButton.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.adjustSizesButton.UseVisualStyleBackColor = true;
            // 
            // checkSizeAccordanceButton
            // 
            this.checkSizeAccordanceButton.Location = new System.Drawing.Point(12, 61);
            this.checkSizeAccordanceButton.Name = "checkSizeAccordanceButton";
            this.checkSizeAccordanceButton.Size = new System.Drawing.Size(187, 25);
            this.checkSizeAccordanceButton.TabIndex = 8;
            this.checkSizeAccordanceButton.Text = "Check Size Accordance";
            this.checkSizeAccordanceButton.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.checkSizeAccordanceButton.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(209, 287);
            this.Controls.Add(this.checkSizeAccordanceButton);
            this.Controls.Add(this.adjustSizesButton);
            this.Controls.Add(this.setupCorpusButton);
            this.Controls.Add(this.alignWithShaftButton);
            this.Controls.Add(this.alignWithHorizontal);
            this.Controls.Add(this.processSelectionButton);
            this.Controls.Add(this.addMateButton);
            this.Controls.Add(this.runSwButton);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(225, 325);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(225, 325);
            this.Name = "MainForm";
            this.Text = "Funny Macros";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button runSwButton;
        private System.Windows.Forms.Button addMateButton;
        private System.Windows.Forms.Button processSelectionButton;
        private System.Windows.Forms.Button alignWithHorizontal;
        private System.Windows.Forms.Button alignWithShaftButton;
        private System.Windows.Forms.Button setupCorpusButton;
        private System.Windows.Forms.Button adjustSizesButton;
        private System.Windows.Forms.Button checkSizeAccordanceButton;
    }
}

