/*
 * By David Barrett, Microsoft Ltd. 2016. Use at your own risk.  No warranties are given.
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

using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System;
using System.Reflection;

namespace InspectorEvents {

    internal delegate void InspectorWrapperClosedEventHandler(Guid id);

    internal abstract class InspectorWrapper {

        public event InspectorWrapperClosedEventHandler Closed;
        public Guid Id { get; private set; }
        public Outlook.Inspector Inspector { get; private set; }

        public InspectorWrapper(Outlook.Inspector inspector) {
            Id = Guid.NewGuid();
            Inspector = inspector;
            // register for Inspector events
            Inspector.AttachmentSelectionChange += Inspector_AttachmentSelectionChange;
            Inspector.BeforeMaximize += Inspector_BeforeMaximize;
            Inspector.BeforeMinimize += Inspector_BeforeMinimize;
            Inspector.BeforeMove += Inspector_BeforeMove;
            Inspector.BeforeSize += Inspector_BeforeSize;
            Inspector.Deactivate += Inspector_Deactivate;
            Inspector.PageChange += Inspector_PageChange;
            ((Outlook.InspectorEvents_10_Event)Inspector).Close += new Outlook.InspectorEvents_10_CloseEventHandler(Inspector_Close);
            ((Outlook.InspectorEvents_10_Event)Inspector).Activate += new Outlook.InspectorEvents_10_ActivateEventHandler(Activate);

            Initialize();
        }

        void Inspector_PageChange(ref string ActivePageName)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Inspector_PageChange");
        }

        void Inspector_Deactivate()
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Inspector_Deactivate");
        }

        void Inspector_BeforeSize(ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Inspector_BeforeSize");
        }

        void Inspector_BeforeMove(ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Inspector_BeforeMove");
        }

        void Inspector_BeforeMinimize(ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Inspector_BeforeMinimize");
        }

        void Inspector_BeforeMaximize(ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Inspector_BeforeMaximize");
        }

        void Inspector_AttachmentSelectionChange()
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Inspector_AttachmentSelectionChange");
        }

        private void Inspector_Close() {
            // unregister Inspector events
            Inspector.AttachmentSelectionChange -= Inspector_AttachmentSelectionChange;
            Inspector.BeforeMaximize -= Inspector_BeforeMaximize;
            Inspector.BeforeMinimize -= Inspector_BeforeMinimize;
            Inspector.BeforeMove -= Inspector_BeforeMove;
            Inspector.BeforeSize -= Inspector_BeforeSize;
            Inspector.Deactivate -= Inspector_Deactivate;
            Inspector.PageChange -= Inspector_PageChange;
            ((Outlook.InspectorEvents_10_Event)Inspector).Close -= new Outlook.InspectorEvents_10_CloseEventHandler(Inspector_Close);
            ((Outlook.InspectorEvents_10_Event)Inspector).Activate -= new Outlook.InspectorEvents_10_ActivateEventHandler(Activate);

            Close();

            Inspector = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            if (Closed != null) Closed(Id);
        }

        protected virtual void Initialize() { }


        protected virtual void Activate()
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Inspector_Activate");
        }

        protected virtual void Close()
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Inspector_Close");
        }

        public static InspectorWrapper GetWrapperFor(Outlook.Inspector inspector)
        {
            string messageClass = (string)inspector.CurrentItem.GetType().InvokeMember("MessageClass", BindingFlags.GetProperty, null, inspector.CurrentItem, null);

            // Create a wrapper depending upon the message class
            switch (messageClass) {
                case "IPM.Contact":
                    return new ContactItemWrapper(inspector);
                case "IPM.Journal":
                    return new JournalItemWrapper(inspector);
                case "IPM.Note":
                    return new MailItemInspectorWrapper(inspector);
                case "IPM.Post":
                    return new PostItemWrapper(inspector);
                case "IPM.Task":
                    return new TaskItemWrapper(inspector);
            }

            // Create a wrapper based on Outlook type
            if (inspector.CurrentItem is Outlook.AppointmentItem) {
                return new AppointmentItemWrapper(inspector);
            }

            // no wrapper found
            return null;
        }
    }
}
