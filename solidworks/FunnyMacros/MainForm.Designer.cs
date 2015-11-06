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
            this.SuspendLayout();
            // 
            // runSwButton
            // 
            this.runSwButton.Location = new System.Drawing.Point(12, 12);
            this.runSwButton.Name = "runSwButton";
            this.runSwButton.Size = new System.Drawing.Size(180, 43);
            this.runSwButton.TabIndex = 0;
            this.runSwButton.Text = "Run SolidWorks.\r\nLoad Components.";
            this.runSwButton.UseVisualStyleBackColor = true;
            // 
            // addMateButton
            // 
            this.addMateButton.Location = new System.Drawing.Point(12, 110);
            this.addMateButton.Name = "addMateButton";
            this.addMateButton.Size = new System.Drawing.Size(180, 43);
            this.addMateButton.TabIndex = 1;
            this.addMateButton.Text = "Add Mate";
            this.addMateButton.UseVisualStyleBackColor = true;
            // 
            // processSelectionButton
            // 
            this.processSelectionButton.Location = new System.Drawing.Point(12, 61);
            this.processSelectionButton.Name = "processSelectionButton";
            this.processSelectionButton.Size = new System.Drawing.Size(180, 43);
            this.processSelectionButton.TabIndex = 3;
            this.processSelectionButton.Text = "Process selection";
            this.processSelectionButton.UseVisualStyleBackColor = true;
            // 
            // alignWithHorizontal
            // 
            this.alignWithHorizontal.Location = new System.Drawing.Point(12, 159);
            this.alignWithHorizontal.Name = "alignWithHorizontal";
            this.alignWithHorizontal.Size = new System.Drawing.Size(180, 43);
            this.alignWithHorizontal.TabIndex = 4;
            this.alignWithHorizontal.Text = "Align With Horizont";
            this.alignWithHorizontal.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(204, 261);
            this.Controls.Add(this.alignWithHorizontal);
            this.Controls.Add(this.processSelectionButton);
            this.Controls.Add(this.addMateButton);
            this.Controls.Add(this.runSwButton);
            this.Name = "MainForm";
            this.Text = "Funny Macros";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button runSwButton;
        private System.Windows.Forms.Button addMateButton;
        private System.Windows.Forms.Button processSelectionButton;
        private System.Windows.Forms.Button alignWithHorizontal;
    }
}

