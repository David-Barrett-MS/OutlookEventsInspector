/*
 * By David Barrett, Microsoft Ltd. 2016-2020. Use at your own risk.  No warranties are given.
 * 
 * DISCLAIMER:
 * THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
 * MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
 * A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
 * MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
 * BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
 * SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
 * OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace InspectorEvents
{
    [ComVisible(true)]
    public class RibbonContextMenu : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public RibbonContextMenu()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("InspectorEvents.RibbonContextMenu.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            Globals.ThisAddIn.FolderContextMenu = ribbonUI;
        }

        public bool IsMonitorFolderVisible(Office.IRibbonControl control)
        {
            try
            {
                Outlook.Folder folder = ((Microsoft.Office.Interop.Outlook.Folder)control.Context);
                if (folder != null)
                    return !FolderWrapper.FolderIsMonitored(folder);
            }
            catch { }
            return false;
        }

        public bool IsUnmonitorFolderVisible(Office.IRibbonControl control)
        {
            try
            {
                Outlook.Folder folder = ((Microsoft.Office.Interop.Outlook.Folder)control.Context);
                if (folder != null)
                    return FolderWrapper.FolderIsMonitored(folder);
            }
            catch { }
            return false;
        }

        public void MonitorFolderItemEvents(Office.IRibbonControl control)
        {
            // Hook into the currently selected folder, and register for events
            try
            {
                Outlook.Folder folder = ((Microsoft.Office.Interop.Outlook.Folder)control.Context);
                FolderWrapper.CreateFolderWrapper(folder);
            }
            catch { }
        }

        public void MonitorFolderItems(Office.IRibbonControl control)
        {
            // Open folder item monitor window targetting the current folder
            try
            {
                Outlook.Folder folder = ((Microsoft.Office.Interop.Outlook.Folder)control.Context);
                FormFolderItemWatcher form = new FormFolderItemWatcher(folder);
                form.Show();
            }
            catch { }
        }

        public void MonitorFolderItemsEWS(Office.IRibbonControl control)
        {
            // Open EWS folder item monitor window targetting the current folder
            try
            {
                Outlook.Folder folder = ((Microsoft.Office.Interop.Outlook.Folder)control.Context);
                FormEWSFolderItemWatcher form = new FormEWSFolderItemWatcher(folder);
                form.Show();
            }
            catch { }
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
