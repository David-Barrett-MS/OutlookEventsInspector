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
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace InspectorEvents
{
    internal delegate void MailItemEventWrapperClosedEventHandler(Guid id);

    class MailItemEventWrapper
    {
        public Outlook.MailItem Item { get; private set; }
        public Guid Id { get; private set; }
        public event MailItemEventWrapperClosedEventHandler Closed;
        private bool _itemIsLoaded = false;
        private bool _analyseWhenLoaded = false;

        public MailItemEventWrapper(Outlook.MailItem item, bool AnalyseWhenLoaded=false)
        {
            Item = item;
            _analyseWhenLoaded = AnalyseWhenLoaded;
            Id = new Guid();
            Item.BeforeRead += new Outlook.ItemEvents_10_BeforeReadEventHandler(Item_BeforeRead);
            Item.Read += new Outlook.ItemEvents_10_ReadEventHandler(Item_Read);
            Item.Unload += new Outlook.ItemEvents_10_UnloadEventHandler(Item_Unload);
            Item.AfterWrite += Item_AfterWrite;
            Item.AttachmentAdd += Item_AttachmentAdd;
            Item.AttachmentRead += Item_AttachmentRead;
            Item.AttachmentRemove += Item_AttachmentRemove;
            Item.BeforeAttachmentAdd += Item_BeforeAttachmentAdd;
            Item.BeforeAttachmentPreview += Item_BeforeAttachmentPreview;
            Item.BeforeAttachmentRead += Item_BeforeAttachmentRead;
            Item.BeforeAttachmentSave += Item_BeforeAttachmentSave;
            Item.BeforeAttachmentWriteToTempFile += Item_BeforeAttachmentWriteToTempFile;
            Item.BeforeAutoSave += Item_BeforeAutoSave;
            Item.BeforeCheckNames += Item_BeforeCheckNames;
            Item.BeforeDelete += Item_BeforeDelete;
            Item.CustomAction += Item_CustomAction;
            Item.CustomPropertyChange += Item_CustomPropertyChange;
            Item.Open += Item_Open;
            Item.PropertyChange += Item_PropertyChange;
            Item.ReadComplete += Item_ReadComplete;
            Item.Write += Item_Write;
        }

        void Item_Write(ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Item_Write (MailItem)");
        }

        void Item_ReadComplete(ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Item_ReadComplete (MailItem)");
            _itemIsLoaded = true;
        }

        void Item_PropertyChange(string Name)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Item_PropertyChange (MailItem)");
        }

        void Item_Open(ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Item_Open (MailItem)");
        }

        void Item_CustomPropertyChange(string Name)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Item_CustomPropertyChange (MailItem)");
        }

        void Item_CustomAction(object Action, object Response, ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Item_CustomAction (MailItem)");
        }

        void Item_BeforeDelete(object Item, ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Item_BeforeDelete (MailItem)");
        }

        void Item_BeforeCheckNames(ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Item_BeforeCheckNames (MailItem)");
        }

        void Item_BeforeAutoSave(ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Item_BeforeAutoSave (MailItem)");
        }

        void Item_BeforeAttachmentWriteToTempFile(Outlook.Attachment Attachment, ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Item_BeforeAttachmentWriteToTempFile (MailItem)");
        }

        void Item_BeforeAttachmentSave(Outlook.Attachment Attachment, ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Item_BeforeAttachmentSave (MailItem)");
        }

        void Item_BeforeAttachmentRead(Outlook.Attachment Attachment, ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Item_BeforeAttachmentRead (MailItem)");
        }

        void Item_BeforeAttachmentPreview(Outlook.Attachment Attachment, ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Item_BeforeAttachmentPreview (MailItem)");
        }

        void Item_BeforeAttachmentAdd(Outlook.Attachment Attachment, ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Item_BeforeAttachmentAdd (MailItem)");
        }

        void Item_AttachmentRemove(Outlook.Attachment Attachment)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Item_AttachmentRemove (MailItem)");
        }

        void Item_AttachmentRead(Outlook.Attachment Attachment)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Item_AttachmentRead (MailItem)");
        }

        void Item_AttachmentAdd(Outlook.Attachment Attachment)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Item_AttachmentAdd (MailItem)");
        }

        void Item_AfterWrite()
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Item_AfterWrite (MailItem)");
        }

        void Item_Read()
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Item_Read (MailItem)");
            if (_analyseWhenLoaded)
            {
                // This is an Item_Read event after we had an Explorer SelectionChange event.
                // This should imply that we are loading the item into the preview pane.  It might be necessary
                // to add more checks here (i.e. to confirm that the item being loaded is actually being shown in the reading pane)

                _analyseWhenLoaded = false; // We only need to do this once
                Globals.ThisAddIn.ReadItemProperty(Item, Globals.ThisAddIn.EventTrackerForm.textBoxPropId.Text);
            }
        }

        void Item_Unload()
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Item_Unload (MailItem)");
            Item.BeforeRead -= new Outlook.ItemEvents_10_BeforeReadEventHandler(Item_BeforeRead);
            Item.Read -= new Outlook.ItemEvents_10_ReadEventHandler(Item_Read);
            Item.Unload -= new Outlook.ItemEvents_10_UnloadEventHandler(Item_Unload);
            Item = null;
            GC.Collect();
            if (Closed != null) Closed(Id);
        }

        void Item_BeforeRead()
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Item_BeforeRead (MailItem)");
        }

        bool itemIsSMIMEEncrypted()
        {
            // Return true if this mail item is SMIME encrypted

            if (!_itemIsLoaded)
            {
                Globals.ThisAddIn.EventTrackerForm.AddLog("Cannot determine SMIME status before item is loaded");
                return false;
            }

            //if (Item.PropertyAccessor)

            return false;
        }
    }
}
