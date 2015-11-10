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

            runButton.Click += (s, e) =>
            {
                ISldWorks solidWorks = Marshal.GetActiveObject("SldWorks.Application") as ISldWorks;
                if (solidWorks != null)
                {
                    new SolidWorksMacro(solidWorks, InputDialog.Show).Run();
                }
                else
                {
                    MessageBox.Show("Please sure Solid Works is running", "Solid Works Not Runnning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            };
        }
    }
}
