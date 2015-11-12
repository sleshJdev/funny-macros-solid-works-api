using System;
using System.Windows.Forms;
using FunnyMacros.Macros;
using System.Drawing;
using SolidWorks.Interop.sldworks;
using System.Runtime.InteropServices;

namespace FunnyMacros.View
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();

            Bitmap bitmap = Bitmap.FromFile(@"resources\icon-app.png") as Bitmap;
            IntPtr pointerIcon = bitmap.GetHicon();
            Icon icon = Icon.FromHandle(pointerIcon);
            Icon = icon;
            icon.Dispose();

            runButton.BackColor = SystemColors.ActiveCaption;
            runButton.FlatAppearance.MouseOverBackColor = runButton.BackColor;
            runButton.BackColorChanged += (s, e) =>
            {
                runButton.FlatAppearance.MouseOverBackColor = runButton.BackColor;
            };
            runButton.Click += (s, e) =>
            {
                Color normal = SystemColors.ActiveCaption;
                Color success = ColorTranslator.FromHtml("#4F8A10");
                Color error = ColorTranslator.FromHtml("#D8000C");
                SolidWorksMacro macro = null;
                try
                {
                    object activeObject = Marshal.GetActiveObject("SldWorks.Application");
                    if (activeObject != null)
                    {
                        ISldWorks solidWorks = activeObject as ISldWorks;
                        if (solidWorks != null)
                        {
                            runButton.BackColor = success;
                            macro = new SolidWorksMacro(solidWorks, InputDialog.Show);
                            if (macro.Run())
                            {
                                runButton.BackColor = normal;
                            }
                            else
                            {
                                runButton.BackColor = error;
                            }
                        }
                    }
                }
                catch
                {
                }
                if (macro == null)
                {
                    runButton.BackColor = error;
                    MessageBox.Show("Please make sure Solid Works has been started", "Solid Works Not Runnning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            };
        }
    }
}
