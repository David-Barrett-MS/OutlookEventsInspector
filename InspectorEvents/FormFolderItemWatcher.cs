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

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace InspectorEvents
{
    public partial class FormFolderItemWatcher : Form
    {
        private Outlook.Folder _watchedFolder = null;
        private Outlook.Items _watchedFolderItems = null;
        public static readonly string _propItemProcessedDate = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/FormFolderItemWatcherProcessed";
        private List<string> _suppressNextChangeEventForProcessedItems;

        public FormFolderItemWatcher(Outlook.Folder folder)
        {
            InitializeComponent();
            _suppressNextChangeEventForProcessedItems = new List<string>();

            // Hook into the folder and watch items
            _watchedFolder = folder;
            _watchedFolderItems = folder.Items;
            _watchedFolderItems.ItemAdd += _watchedFolderItems_ItemAdd;
            _watchedFolderItems.ItemChange += _watchedFolderItems_ItemChange;
            _watchedFolderItems.ItemRemove += _watchedFolderItems_ItemRemove;
            AddLog(String.Format("Hooked into folder events.  Currently {0} items in folder.", folder.Items.Count));
            this.Text = String.Format("Folder Item Watcher: {0}", folder.Name);
            this.TopMost = checkBoxAlwaysOnTop.Checked;

            // Check for any items that we missed while not watching (e.g. Outlook closed
            ProcessMissedItems();
        }

        ~FormFolderItemWatcher()
        {
            try
            {
                _watchedFolderItems.ItemAdd -= _watchedFolderItems_ItemAdd;
                _watchedFolderItems.ItemChange -= _watchedFolderItems_ItemChange;
                _watchedFolderItems.ItemRemove -= _watchedFolderItems_ItemRemove;
            }
            catch { }
        }

        private void AddLog(string Log)
        {
            if (listBoxFolderItems.Items.Count > 1000)
                listBoxFolderItems.Items.RemoveAt(0);
            Log = String.Format("{0:HH:mm:ss.ff}  {1}", DateTime.Now, Log);
            listBoxFolderItems.Items.Add(Log);
            listBoxFolderItems.SelectedIndex = listBoxFolderItems.Items.Count - 1;
        }

        void _watchedFolderItems_ItemRemove()
        {
            AddLog("Item removed");
        }

        void _watchedFolderItems_ItemChange(object Item)
        {
            if (!(Item is Outlook.MailItem))
            {
                AddLog("Item changed: not a MailItem (ignored)");
                return;
            }

            if (_suppressNextChangeEventForProcessedItems.Contains((Item as Outlook.MailItem).EntryID))
            {
                _suppressNextChangeEventForProcessedItems.Remove((Item as Outlook.MailItem).EntryID);
                return;
            }
            AddLog(String.Format("Item changed: {0}", (Item as Outlook.MailItem).Subject));
        }

        void _watchedFolderItems_ItemAdd(object Item)
        {
            ProcessItem(Item);
        }


        private void checkBoxAlwaysOnTop_CheckedChanged(object sender, EventArgs e)
        {
            this.TopMost = checkBoxAlwaysOnTop.Checked;
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void ProcessItem(object OutlookItem)
        {
            if (OutlookItem is Outlook.MailItem)
            {
                MarkAsProcessed((Outlook.MailItem)OutlookItem);
                AddLog(String.Format("Item added: {0}", (OutlookItem as Outlook.MailItem).Subject));
                return;
            }

            //AddLog(String.Format("Item added: {0}", (Item as Outlook.MailItem).Subject));
        }

        private void MarkAsProcessed(Outlook.MailItem mailItem)
        {
            // Add custom property so that we know this item has been processed
            try
            {
                mailItem.PropertyAccessor.SetProperty(_propItemProcessedDate, String.Format("P-{0}", DateTime.Now.ToUniversalTime().ToString()));
                // We log the EntryID of this item so that we can suppress the next change event (which is caused by us writing this property)
                _suppressNextChangeEventForProcessedItems.Add(mailItem.EntryID);  
                mailItem.Save();
            }
            catch { }
        }

        private bool HasItemBeenProcessed(Outlook.MailItem mailItem)
        {
            // Check for custom property and determine whether or not we've processed this item already
            try
            {
                DateTime processedOn = mailItem.PropertyAccessor.GetProperty(_propItemProcessedDate);
                if (processedOn < DateTime.Now)
                    return true;
            }
                // We'll get an error if the property doesn't exist, or we can't read it
            catch { }

            return false;
        }

        private void ProcessMissedItems()
        {
            // We process all items in the folder that we've missed (e.g. Outlook was closed, or too many changes happened at once)
            // We do this by looking for any items that do not have our custom property set

            Outlook.Items missedItems = _watchedFolder.Items;

            // As we're filtering on a custom property that we haven't added to the folder, we need to add the
            // property type (0x0000001f) to the end of the property definition (otherwise, our filter won't work)
            string filter = String.Format("@SQL=NOT \"{0}/0x0000001f\" LIKE 'P-%'", _propItemProcessedDate);
            //AddLog("Filter: " + filter);
            object missedItem = missedItems.Find(filter);
            if (missedItem == null)
                return;

            ProcessItem(missedItem);
            while (missedItem != null)
            {
                missedItem = null;
                missedItem = missedItems.FindNext();
                if (missedItem != null)
                {
                    ProcessItem(missedItem);
                }
            }
        }
    }
}
