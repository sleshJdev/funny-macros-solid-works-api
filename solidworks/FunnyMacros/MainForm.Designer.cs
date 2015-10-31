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
            this.loadShaftButton = new System.Windows.Forms.Button();
            this.processSelectionButton = new System.Windows.Forms.Button();
            this.loadLocatorsButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // runSwButton
            // 
            this.runSwButton.Location = new System.Drawing.Point(12, 12);
            this.runSwButton.Name = "runSwButton";
            this.runSwButton.Size = new System.Drawing.Size(149, 43);
            this.runSwButton.TabIndex = 0;
            this.runSwButton.Text = "Run SolidWorks.\r\nLoad Components.";
            this.runSwButton.UseVisualStyleBackColor = true;
            // 
            // addMateButton
            // 
            this.addMateButton.Location = new System.Drawing.Point(167, 61);
            this.addMateButton.Name = "addMateButton";
            this.addMateButton.Size = new System.Drawing.Size(123, 43);
            this.addMateButton.TabIndex = 1;
            this.addMateButton.Text = "Add Mate";
            this.addMateButton.UseVisualStyleBackColor = true;
            // 
            // loadShaftButton
            // 
            this.loadShaftButton.Location = new System.Drawing.Point(12, 187);
            this.loadShaftButton.Name = "loadShaftButton";
            this.loadShaftButton.Size = new System.Drawing.Size(149, 23);
            this.loadShaftButton.TabIndex = 2;
            this.loadShaftButton.Text = "Load Shaft";
            this.loadShaftButton.UseVisualStyleBackColor = true;
            // 
            // processSelectionButton
            // 
            this.processSelectionButton.Location = new System.Drawing.Point(167, 12);
            this.processSelectionButton.Name = "processSelectionButton";
            this.processSelectionButton.Size = new System.Drawing.Size(123, 43);
            this.processSelectionButton.TabIndex = 3;
            this.processSelectionButton.Text = "Process selection";
            this.processSelectionButton.UseVisualStyleBackColor = true;
            // 
            // loadLocatorsButton
            // 
            this.loadLocatorsButton.Location = new System.Drawing.Point(297, 13);
            this.loadLocatorsButton.Name = "loadLocatorsButton";
            this.loadLocatorsButton.Size = new System.Drawing.Size(114, 42);
            this.loadLocatorsButton.TabIndex = 4;
            this.loadLocatorsButton.Text = "Load locators";
            this.loadLocatorsButton.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(493, 261);
            this.Controls.Add(this.loadLocatorsButton);
            this.Controls.Add(this.processSelectionButton);
            this.Controls.Add(this.loadShaftButton);
            this.Controls.Add(this.addMateButton);
            this.Controls.Add(this.runSwButton);
            this.Name = "MainForm";
            this.Text = "Funny Macros";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button runSwButton;
        private System.Windows.Forms.Button addMateButton;
        private System.Windows.Forms.Button loadShaftButton;
        private System.Windows.Forms.Button processSelectionButton;
        private System.Windows.Forms.Button loadLocatorsButton;
    }
}

