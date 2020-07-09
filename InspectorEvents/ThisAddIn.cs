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
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;


namespace InspectorEvents
{
    public partial class ThisAddIn
    {
        private Outlook.Inspectors _inspectors;
        private Dictionary<Guid, MailItemEventWrapper> _mailItems;
        //private Outlook.Explorer _activeExplorer;
        private Outlook.Application _application;
        private List<Outlook.MailItem> _trackedMailItems;
        private Office.IRibbonUI _folderContextMenu;
        private Dictionary<Guid, InspectorWrapper> _wrappedInspectors;
        private Dictionary<Guid, ExplorerWrapper> _wrappedExplorers;

        public bool ExplorerSelectionChangeOccurred { get; set; }
        public FormEventTracker EventTrackerForm { get; private set; }

        public Office.IRibbonUI FolderContextMenu
        {
            get { return _folderContextMenu; }
            set { _folderContextMenu = value; }
        }

        public void ReadItemProperty(object Item, string PropId)
        {
            if (Item is Outlook.MailItem)
            {
                try
                {
                    Globals.ThisAddIn.EventTrackerForm.AddLog($"Requesting property {PropId}");
                    object propVal = ((Outlook.MailItem)Item).PropertyAccessor.GetProperty($"http://schemas.microsoft.com/mapi/proptag/{PropId}");
                    Globals.ThisAddIn.EventTrackerForm.AddLog($"Property {PropId} retrieved successfully: {propVal.ToString()}");
                }
                catch (System.Exception ex)
                {
                    Globals.ThisAddIn.EventTrackerForm.AddLog($"Failed to retrieve property {PropId}: {ex.Message}");
                }
            }
            else
                Globals.ThisAddIn.EventTrackerForm.AddLog("Item ignored as it is not a MailItem");
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // This method is called by Outlook for the add-in to initialise, so we need to hook into events here
            _application = Globals.ThisAddIn.Application;
            _wrappedExplorers = new Dictionary<Guid, ExplorerWrapper>();
            _wrappedInspectors = new Dictionary<Guid, InspectorWrapper>();
            _inspectors = _application.Inspectors;
            _inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(WrapInspector);

            // Wrap any existing Explorers
            foreach (Outlook.Explorer explorer in _application.Explorers)
            {
                WrapExplorer(explorer);
            }

            // Wrap any existing Inspectors
            foreach (Outlook.Inspector inspector in _inspectors) {
                WrapInspector(inspector);
            }

            _trackedMailItems = new List<Outlook.MailItem>();
            //HookActiveExplorer();

            _mailItems = new Dictionary<Guid, MailItemEventWrapper>();
            _application.ItemLoad += new Outlook.ApplicationEvents_11_ItemLoadEventHandler(_application_ItemLoad);
            EventTrackerForm = new FormEventTracker();
            EventTrackerForm.Show();
        }

        void invalidateFolderContextMenu()
        {
            try
            {
                if (_folderContextMenu != null)
                {
                    _folderContextMenu.InvalidateControl("FolderMonitorItemEvents");
                    _folderContextMenu.InvalidateControl("FolderUnmonitorItemEvents");
                }
            }
            catch { }
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonContextMenu();
        }

        void _application_ItemLoad(object Item)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Application_ItemLoad");
            if (!(Item is Outlook.MailItem))
                return;

            Outlook.MailItem mailItem = Item as Outlook.MailItem;
            MailItemEventWrapper wrapper = new MailItemEventWrapper(mailItem, ExplorerSelectionChangeOccurred);

            wrapper.Closed += new MailItemEventWrapperClosedEventHandler(mailItemEventWrapper_Closed);
            _mailItems[wrapper.Id]=wrapper;
        }

        void mailItemEventWrapper_Closed(Guid id)
        {
            _mailItems.Remove(id);
        }

        void TrackSelectedItems()
        {
        }

        /*        
        void HookActiveExplorerSelection()
        {
            for (int i = _trackedMailItems.Count-1; i>=0; i--)
            {
                Outlook.MailItem oItem = _trackedMailItems[i];
                try
                {
                    ((Outlook.ItemEvents_10_Event)oItem).Forward -= new Outlook.ItemEvents_10_ForwardEventHandler(ThisAddIn_Forward);
                    ((Outlook.ItemEvents_10_Event)oItem).Reply -= new Outlook.ItemEvents_10_ReplyEventHandler(ThisAddIn_Reply);
                    ((Outlook.ItemEvents_10_Event)oItem).ReplyAll -= new Outlook.ItemEvents_10_ReplyAllEventHandler(ThisAddIn_ReplyAll);
                    ((Outlook.ItemEvents_10_Event)oItem).AttachmentRead -= new Outlook.ItemEvents_10_AttachmentReadEventHandler(ThisAddIn_AttachmentRead);
                }
                catch { }
                _trackedMailItems.RemoveAt(i);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem);
            }

            foreach (object oItem in _activeExplorer.Selection)
            {
                if (oItem is Outlook.MailItem)
                {
                    Outlook.MailItem oMailItem = oItem as Outlook.MailItem;
                    if (oMailItem.MessageClass == "IPM.Note")
                    {
                        ((Outlook.ItemEvents_10_Event)oMailItem).Forward += new Outlook.ItemEvents_10_ForwardEventHandler(ThisAddIn_Forward);
                        ((Outlook.ItemEvents_10_Event)oMailItem).Reply += new Outlook.ItemEvents_10_ReplyEventHandler(ThisAddIn_Reply);
                        ((Outlook.ItemEvents_10_Event)oMailItem).ReplyAll += new Outlook.ItemEvents_10_ReplyAllEventHandler(ThisAddIn_ReplyAll);
                        ((Outlook.ItemEvents_10_Event)oMailItem).AttachmentRead += new Outlook.ItemEvents_10_AttachmentReadEventHandler(ThisAddIn_AttachmentRead);
                        _trackedMailItems.Add(oMailItem);
                    }
                }
            }
        }
        */

        void ThisAddIn_AttachmentRead(Outlook.Attachment Attachment)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("ThisAddIn_AttachmentRead");
        }

        void ThisAddIn_ReplyAll(object Response, ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("ThisAddIn_ReplyAll");
        }

        void ThisAddIn_Reply(object Response, ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("ThisAddIn_Reply");
        }

        void ThisAddIn_Forward(object Forward, ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("ThisAddIn_Forward");
        }

        void WrapInspector(Outlook.Inspector inspector)
        {
            InspectorWrapper wrapper = InspectorWrapper.GetWrapperFor(inspector);
            if (wrapper != null) {
                wrapper.Closed += new InspectorWrapperClosedEventHandler(inspectorWrapper_Closed);
                _wrappedInspectors[wrapper.Id] = wrapper;
            }
        }

        void WrapExplorer(Outlook.Explorer explorer)
        {
            ExplorerWrapper wrapper = new ExplorerWrapper(explorer);
            if (wrapper != null)
            {
                wrapper.Closed += explorerWrapper_Closed;
            }
        }

        private void explorerWrapper_Closed(Guid id)
        {
            _wrappedExplorers.Remove(id);
        }

        void inspectorWrapper_Closed(Guid id) {
            _wrappedInspectors.Remove(id); 
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Clean-up
            _wrappedInspectors.Clear();
            _inspectors.NewInspector -=new Outlook.InspectorsEvents_NewInspectorEventHandler(WrapInspector); 
            _inspectors = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}


