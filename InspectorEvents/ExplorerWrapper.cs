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
using Microsoft.Win32;

namespace InspectorEvents

{
    internal delegate void ExplorerEventWrapperClosedEventHandler(Guid id);

    class ExplorerWrapper
    {
        public Outlook.Explorer Explorer { get; private set; }
        public event MailItemEventWrapperClosedEventHandler Closed;
        public Guid Id { get; private set; }


        public ExplorerWrapper(Outlook.Explorer explorer)
        {
            Explorer = explorer;
            Id = new Guid();
            explorer.AttachmentSelectionChange += Explorer_AttachmentSelectionChange;
            explorer.BeforeFolderSwitch += Explorer_BeforeFolderSwitch;
            explorer.BeforeItemCopy += Explorer_BeforeItemCopy;
            explorer.BeforeItemCut += Explorer_BeforeItemCut;
            explorer.BeforeItemPaste += Explorer_BeforeItemPaste;
            explorer.BeforeMaximize += Explorer_BeforeMaximize;
            explorer.BeforeMinimize += Explorer_BeforeMinimize;
            explorer.BeforeMove += Explorer_BeforeMove;
            explorer.BeforeSize += Explorer_BeforeSize;
            explorer.BeforeViewSwitch += Explorer_BeforeViewSwitch;
            explorer.Deactivate += Explorer_Deactivate;
            explorer.FolderSwitch += Explorer_FolderSwitch;
            explorer.InlineResponse += Explorer_InlineResponse;
            explorer.InlineResponseClose += Explorer_InlineResponseClose;
            explorer.SelectionChange += Explorer_SelectionChange;
            explorer.ViewSwitch += Explorer_ViewSwitch;
        }

        private void Explorer_ViewSwitch()
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Explorer_ViewSwitch");
        }

        private void Explorer_SelectionChange()
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Explorer_SelectionChange");

            if (Explorer.Selection.Count == 0)
            {
                Globals.ThisAddIn.EventTrackerForm.AddLog($"No items currently selected");
                return;
            }

            if (Globals.ThisAddIn.EventTrackerForm.checkBoxRetrievePropertyonSelectionChange.Checked)
            {
                if (encryptedMessagePreviewDisabled() && Globals.ThisAddIn.EventTrackerForm.checkBoxRetrievePropOnlyIfVisible.Checked)
                {
                    // Encrypted message preview is disabled, so we want to ensure we don't trigger loading of an encrypted message
                    // At this point, we don't know what the item is, and reading a property to find out will trigger it to be 
                    // loaded and decrypted.  So, here we simply set a flag to check in the next ItemLoad event for the property we want
                    // (ItemLoad will not fire for encrypted messages, hence avoiding the issue of decrypting automatically)

                    Globals.ThisAddIn.ExplorerSelectionChangeOccurred = true;
                    Globals.ThisAddIn.EventTrackerForm.AddLog($"Delaying item property read until ItemRead (which may not trigger, dependent upon settings)");
                    return;
                }

                // As encrypted message preview is not disabled (or we are ignoring it), we can read the property from the item now (it will be loaded anyway)

                if (Explorer.Selection.Count > 1)
                    Globals.ThisAddIn.EventTrackerForm.AddLog($"Attempting to read property from first item in selection");

                Globals.ThisAddIn.ReadItemProperty(Explorer.Selection[1], Globals.ThisAddIn.EventTrackerForm.textBoxPropId.Text);
            }
        }

        private void Explorer_InlineResponseClose()
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Explorer_InlineResponseClose");
        }

        private void Explorer_InlineResponse(object Item)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Explorer_InlineResponse");
        }

        private void Explorer_FolderSwitch()
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Explorer_FolderSwitch");
        }

        private void Explorer_Deactivate()
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Explorer_Deactivate");
        }

        private void Explorer_BeforeViewSwitch(object NewView, ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Explorer_BeforeViewSwitch");
        }

        private void Explorer_BeforeSize(ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Explorer_BeforeSize");
        }

        private void Explorer_BeforeMove(ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Explorer_BeforeMove");
        }

        private void Explorer_BeforeMinimize(ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Explorer_BeforeMinimize");
        }

        private void Explorer_BeforeMaximize(ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Explorer_BeforeMaximize");
        }

        private void Explorer_BeforeItemPaste(ref object ClipboardContent, Outlook.MAPIFolder Target, ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Explorer_BeforeItemPaste");
        }

        private void Explorer_BeforeItemCut(ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Explorer_BeforeItemCut");
        }

        private void Explorer_BeforeItemCopy(ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Explorer_BeforeItemCopy");
        }

        private void Explorer_BeforeFolderSwitch(object NewFolder, ref bool Cancel)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Explorer_BeforeFolderSwitch");
        }

        private void Explorer_AttachmentSelectionChange()
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog("Explorer_AttachmentSelectionChange");
        }

        bool encryptedMessagePreviewDisabled()
        {
            // Read registry key to find out if encrypted message preview is disabled
            // Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Security\DisableEncryptedMessagePreview=DWORD:1

            bool previewDisabled = false;
            RegistryKey sk = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Office\\16.0\\Outlook\\Security");
            if (sk != null)
            {
                previewDisabled = ((int)sk.GetValue("DisableEncryptedMessagePreview") == 1);
                sk.Close();
            }
            Globals.ThisAddIn.EventTrackerForm.AddLog($"DisableEncryptedMessagePreview: {previewDisabled}");
            return previewDisabled;
        }
    }
}
