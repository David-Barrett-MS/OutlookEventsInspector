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
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace InspectorEvents
{
    class FolderWrapper
    {
        static Dictionary<String, FolderWrapper> _watchedFolders = new Dictionary<String, FolderWrapper>();
        Outlook.Folder _thisFolder = null;
        Outlook.Items _thisFolderItems = null;

        public static FolderWrapper CreateFolderWrapper(Outlook.Folder folder)
        {
            if (_watchedFolders.ContainsKey(folder.EntryID))
            {
                // We're already monitoring this folder
                return _watchedFolders[folder.EntryID];
            }

            return new FolderWrapper(folder);
        }

        public static bool FolderIsMonitored(Outlook.Folder folder)
        {
            // Are we watching this folder already?
            return _watchedFolders.ContainsKey(folder.EntryID);
        }

        FolderWrapper(Outlook.Folder folder)
        {
            // Create a new wrapper
            _thisFolder = folder;
            _thisFolderItems = folder.Items;
            _watchedFolders.Add(folder.EntryID, this);
            RegisterForEvents();
        }

        ~FolderWrapper()
        {
            // Destructor, remove this wrapper
            if (_watchedFolders.ContainsKey(_thisFolder.EntryID))
            {
                UnregisterEvents();
                _watchedFolders.Remove(_thisFolder.EntryID);
                _thisFolderItems = null;
                _thisFolder = null;
            }
        }

        void RegisterForEvents()
        {
            // Register folder events
            _thisFolder.BeforeFolderMove += _thisFolder_BeforeFolderMove;
            _thisFolder.BeforeItemMove += _thisFolder_BeforeItemMove;

            // Register folder item events
            _thisFolderItems.ItemAdd += Items_ItemAdd;
            _thisFolderItems.ItemChange += Items_ItemChange;
            _thisFolderItems.ItemRemove += Items_ItemRemove;
        }

        void UnregisterEvents()
        {
            // Unregister folder events
            _thisFolder.BeforeFolderMove -= _thisFolder_BeforeFolderMove;
            _thisFolder.BeforeItemMove -= _thisFolder_BeforeItemMove;

            // Unregister folder item events
            _thisFolderItems.ItemAdd -= Items_ItemAdd;
            _thisFolderItems.ItemChange -= Items_ItemChange;
            _thisFolderItems.ItemRemove -= Items_ItemRemove;
        }

        void AddLog(string eventInfo)
        {
            Globals.ThisAddIn.EventTrackerForm.AddLog(String.Format("{0}->{1}", _thisFolder.Name, eventInfo));
        }

        void Items_ItemRemove()
        {
            AddLog("Items_ItemRemove");
        }

        void Items_ItemChange(object Item)
        {
            AddLog("Items_ItemChange");
        }

        void Items_ItemAdd(object Item)
        {
            AddLog("Items_ItemAdd");
        }

        void _thisFolder_BeforeItemMove(object Item, Outlook.MAPIFolder MoveTo, ref bool Cancel)
        {
            AddLog("_BeforeItemMove");
        }

        void _thisFolder_BeforeFolderMove(Outlook.MAPIFolder MoveTo, ref bool Cancel)
        {
            AddLog("_BeforeFolderMove");
        }
    }
}
