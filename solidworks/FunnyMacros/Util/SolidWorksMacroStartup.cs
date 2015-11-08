using System;
using System.Diagnostics;
using System.Management;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;

namespace FunnyMacros.Util
{
    public partial class SolidWorksMacro
    {
        #region Win32 imports
        [DllImport("ole32.dll")]
        static extern int CreateItemMoniker([MarshalAs(UnmanagedType.LPWStr)] string
           lpszDelim, [MarshalAs(UnmanagedType.LPWStr)] string lpszItem,
           out System.Runtime.InteropServices.ComTypes.IMoniker ppmk);

        [DllImport("ole32.dll")]
        static extern int GetRunningObjectTable(uint reserved,
           out System.Runtime.InteropServices.ComTypes.IRunningObjectTable pprot);
        #endregion

        private void SolidWorksMacro_Startup(object sender, EventArgs e)
        {
            bool proxyFailed = false;
            try
            {
                this.solidWorks = (SldWorks)this.solidWorks;
            }
            catch (Exception)
            {
                proxyFailed = true;
            }

            if (!proxyFailed)
            {
                SolidWorks.RunMacroResult runResult = SolidWorks.RunMacroResult.Run;//RunMacro();

                if (runResult == SolidWorks.RunMacroResult.NoneSpecified)
                {
                    try
                    {
                        //SwMacroSetup();
                    }
                    catch (Exception ex) { System.Diagnostics.Debug.WriteLine(ex.ToString()); }
                    //Main();
                    try
                    {
                        //SwMacroCleanup();
                    }
                    catch (Exception ex) { System.Diagnostics.Debug.WriteLine(ex.ToString()); }
                }

                return;
            }
            try
            {
                Process sldProc = GetParentProcess(Process.GetCurrentProcess());
                while (sldProc != null)
                {
                    if (String.Compare(sldProc.ProcessName, "sldworks", true) == 0)
                    {
                        break;
                    }
                    else
                    {
                        sldProc = GetParentProcess(sldProc);
                    }
                }
                System.Runtime.InteropServices.ComTypes.IRunningObjectTable rot;
                System.Runtime.InteropServices.ComTypes.IMoniker ppmk;
                CreateItemMoniker(null, "SolidWorks_PID_" + sldProc.Id.ToString(), out ppmk);
                GetRunningObjectTable(0, out rot);
                object swAppObj = null;
                rot.GetObject(ppmk, out swAppObj);
                if (swAppObj != null)
                {
                    this.solidWorks = (SldWorks)swAppObj;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            //Main();
        }

        private void SolidWorksMacro_Shutdown(object sender, EventArgs e)
        {

        }

        private void InternalStartup()
        {
            //this.Startup += new System.EventHandler(SolidWorksMacro_Startup);
            //this.Shutdown += new System.EventHandler(SolidWorksMacro_Shutdown);
        }

        Process GetParentProcess(Process childProc)
        {
            int childPID = childProc.Id;

            WqlObjectQuery wqlQuery = new WqlObjectQuery("Select * from Win32_Process where ProcessID=" + childPID.ToString());
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(wqlQuery);
            if (searcher == null)
            {
                return null;
            }
            foreach (ManagementObject disk in searcher.Get())
            {
                int procID = 0;
                if (int.TryParse(disk.GetPropertyValue("ParentProcessId").ToString(), out procID))
                {
                    return Process.GetProcessById(procID);
                }
            }

            return null;
        }
    }
}


