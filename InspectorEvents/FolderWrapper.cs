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
