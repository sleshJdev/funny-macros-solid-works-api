namespace FunnyMacros.View
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.runButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // runButton
            // 
            this.runButton.BackColor = System.Drawing.SystemColors.Control;
            this.runButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.runButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.runButton.FlatAppearance.BorderColor = System.Drawing.SystemColors.Control;
            this.runButton.FlatAppearance.MouseDownBackColor = System.Drawing.SystemColors.Control;
            this.runButton.FlatAppearance.MouseOverBackColor = System.Drawing.SystemColors.Control;
            this.runButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.runButton.Image = ((System.Drawing.Image)(resources.GetObject("runButton.Image")));
            this.runButton.Location = new System.Drawing.Point(-10, -9);
            this.runButton.Name = "runButton";
            this.runButton.Size = new System.Drawing.Size(168, 145);
            this.runButton.TabIndex = 3;
            this.runButton.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.runButton.UseVisualStyleBackColor = false;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(149, 127);
            this.Controls.Add(this.runButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(165, 165);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(165, 165);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Funny Macros";
            this.TopMost = true;
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button runButton;
    }
}

