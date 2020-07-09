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
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Exchange.WebServices.Data;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace InspectorEvents
{
    public partial class FormEWSFolderItemWatcher : Form
    {
        ExchangeService _exchangeService = null;
        string _primarySmtpAddress = string.Empty;
        Outlook.Folder _folder = null;
        FolderId _folderId = null;
        StreamingSubscription _streamingSubscription = null;
        StreamingSubscriptionConnection _streamingSubscriptionConnection = null;

        public FormEWSFolderItemWatcher(Outlook.Folder folder)
        {
            // We need to initialise EWS and subscribe to the folder for notifications
            // Once we've pulled the EWS information from Outlook, we can do any processing that doesn't
            // access the Outlook Object Model on another thread

            InitializeComponent();
            this.TopMost = checkBoxAlwaysOnTop.Checked;
            _folder = folder;

            CreateSubscription();
        }

        private void CreateSubscription()
        {
            if (InitialiseEWS())
            {
                SubscribeForNotifications();
            }
            else
                AddLog("Failed to initialise EWS");
        }

        private bool InitialiseEWS()
        {
            //string autodiscoverXml = Globals.ThisAddIn.Application.Session.AutoDiscoverXml;
            XmlDocument autodiscoverXml = new XmlDocument();
            autodiscoverXml.LoadXml(Globals.ThisAddIn.Application.Session.AutoDiscoverXml);
            _exchangeService = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            _exchangeService.TraceListener = new ClassTraceListener("c:\temp\trace.txt");
            _exchangeService.TraceFlags = TraceFlags.All;
            _exchangeService.TraceEnabled = true;

            // We want the EWS endpoint, and the user's email address
            _primarySmtpAddress = Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress;
            if (String.IsNullOrEmpty(_primarySmtpAddress))
                return false;
            if (!ReadEWSSettings(autodiscoverXml))
                return false;

            return true;
        }

        private bool ReadEWSSettings(XmlDocument AutodiscoverXml)
        {
            // Check the response and extract what we need...
            if (AutodiscoverXml == null) return false;

            // Check for Ews Url
            string sElement = "EwsUrl";
            string sEwsUrl = ReadElement(AutodiscoverXml, sElement);
            if (String.IsNullOrEmpty(sEwsUrl))
            {
                sElement = "ExternalEwsUrl";
                sEwsUrl = ReadUserSetting(AutodiscoverXml, sElement);
                if (String.IsNullOrEmpty(sEwsUrl))
                {
                    sElement = "InternalEwsUrl";
                    sEwsUrl = ReadUserSetting(AutodiscoverXml, sElement);
                }
            }

            if (!String.IsNullOrEmpty(sEwsUrl))
            {
                _exchangeService.Url = new Uri(sEwsUrl);
                AddLog($"Using EWS URL: {sEwsUrl}");
                return true;
            }

            AddLog($"Failed to determine EWS Url");
            return false;
        }

        private string ReadElement(XmlDocument AutodiscoverXml, string ElementName)
        {
            XmlNodeList oEwsUrls = AutodiscoverXml.GetElementsByTagName(ElementName);
            if (oEwsUrls.Count > 0)
            {
                // We have an EWL Url
                XmlNode oEwsUrl = oEwsUrls.Item(0);
                return oEwsUrl.InnerText;
            }
            return null;
        }

        private string ReadUserSetting(XmlDocument AutodiscoverXml, string SettingName)
        {
            XmlNodeList oSettings = AutodiscoverXml.GetElementsByTagName("UserSetting");
            foreach (XmlNode oSetting in oSettings)
            {
                string sName = "";
                string sValue = "";
                foreach (XmlNode oNode in oSetting.ChildNodes)
                {
                    switch (oNode.Name)
                    {
                        case "Name":
                            sName = oNode.InnerText;
                            break;

                        case "Value":
                            sValue = oNode.InnerText;
                            break;
                    }
                }
                if (sName == SettingName)
                {
                    return sValue;
                }
            }
            return null;
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private bool SubscribeForNotifications()
        {
            try
            {
                AlternateId outlookFolderId = new AlternateId(IdFormat.HexEntryId, _folder.EntryID, _primarySmtpAddress, false);
                AlternateId ewsFolderId = _exchangeService.ConvertId(outlookFolderId, IdFormat.EwsId) as AlternateId;
                _folderId = new FolderId(ewsFolderId.UniqueId);
            }
            catch (Exception ex)
            {
                AddLog(String.Format("Failed to obtain EWS Folder Id: {0}", ex.Message));
                if (ex.Message.Contains("Unauthorized") || ex.Message.Contains("401"))
                    AddLog("Currently only Windows auth will work (on-prem only)");
                return false;
            }

            try
            {
                _streamingSubscription = _exchangeService.SubscribeToStreamingNotifications(new FolderId[] { _folderId }, EventType.Created, EventType.Moved, EventType.Copied, EventType.Modified, EventType.NewMail, EventType.Deleted);
                // If we have a watermark, we set this so that we don't miss any events
            }
            catch (Exception ex)
            {
                AddLog(String.Format("Error creating subscription: {0}", ex.Message));
                return false;
            }

            try
            {
                _streamingSubscriptionConnection = new StreamingSubscriptionConnection(_exchangeService, 30);
                _streamingSubscriptionConnection.AddSubscription(_streamingSubscription);
            }
            catch (Exception ex)
            {
                AddLog(String.Format("Error creating subscription connection: {0}", ex.Message));
                return false;
            }

            _streamingSubscriptionConnection.OnNotificationEvent += _streamingSubscriptionConnection_OnNotificationEvent;
            _streamingSubscriptionConnection.OnDisconnect += _streamingSubscriptionConnection_OnDisconnect;
            _streamingSubscriptionConnection.OnSubscriptionError += _streamingSubscriptionConnection_OnSubscriptionError;

            try
            {
                _streamingSubscriptionConnection.Open();
            }
            catch (Exception ex)
            {
                AddLog(String.Format("Error opening subscription connection: {0}", ex.Message));
                return false;
            } 
            
            AddLog("Successfully subscribed for notifications");
            return true;
        }

        private void AddLog(string Log)
        {
            Log = String.Format("{0:HH:mm:ss.ff}  {1}", DateTime.Now, Log);
            if (listBoxFolderItems.InvokeRequired)
            {
                listBoxFolderItems.Invoke(new MethodInvoker(delegate()
                {
                    if (listBoxFolderItems.Items.Count > 1000)
                        listBoxFolderItems.Items.RemoveAt(0);
                    listBoxFolderItems.Items.Add(Log);
                    listBoxFolderItems.SelectedIndex = listBoxFolderItems.Items.Count - 1;
                }));
            }
            else
            {
                if (listBoxFolderItems.Items.Count > 1000)
                    listBoxFolderItems.Items.RemoveAt(0);
                listBoxFolderItems.Items.Add(Log);
                listBoxFolderItems.SelectedIndex = listBoxFolderItems.Items.Count - 1;
            }
        }

        void _streamingSubscriptionConnection_OnSubscriptionError(object sender, SubscriptionErrorEventArgs args)
        {
            AddLog(String.Format("EWS subscription error: {0}", args.Exception));
        }

        void _streamingSubscriptionConnection_OnDisconnect(object sender, SubscriptionErrorEventArgs args)
        {
            // We've lost the subscription. We'll try to recreate it
            AddLog("Subscription disconnection event received");
            CreateSubscription();
        }

        void _streamingSubscriptionConnection_OnNotificationEvent(object sender, NotificationEventArgs args)
        {
            foreach (NotificationEvent notificationEvent in args.Events)
            {
                if (notificationEvent is ItemEvent)
                {
                    ItemEvent itemEvent = notificationEvent as ItemEvent;
                    AddLog(String.Format("EWS Event received: {0}", notificationEvent.EventType.ToString()));
                    ThreadPool.QueueUserWorkItem(new WaitCallback(GetOutlookItemInfo), itemEvent);
                }
            }
        }

        void GetOutlookItemInfo(object e)
        {
            // Get more info for the given item.  This will run on it's own thread
            // so that the main program can continue as usual (we won't hold anything up)
            ItemEvent itemEvent = (ItemEvent)e;

            if (itemEvent.EventType == EventType.Deleted)
            {
                AddLog("Deleted event - no further information available");
                return;
            }

            ExchangeService ewsMoreInfoService = new ExchangeService(_exchangeService.RequestedServerVersion);
            ewsMoreInfoService.Url = _exchangeService.Url;
            ewsMoreInfoService.TraceListener = _exchangeService.TraceListener;
            ewsMoreInfoService.TraceFlags = TraceFlags.All;
            ewsMoreInfoService.TraceEnabled = true;

            try
            {
                // We bind to the actual EWS item, so can read properties from it as we like
                // If we want to do any processing via Outlook, then we need to obtain the EntryId (that is
                // the easiest way to open the item via OOM)
                // Note that we can't do any Outlook OM calls from here as we are on another thread, so we'd
                // need to implement some kind of queue to process items on the main thread.
                // This is beyond the scope of this sample code.

                Item item = Item.Bind(ewsMoreInfoService, itemEvent.ItemId);
                AddLog(String.Format("{0}: {1}", itemEvent.EventType.ToString(), item.Subject));
            }
            catch { }
        }

        private void checkBoxAlwaysOnTop_CheckedChanged(object sender, EventArgs e)
        {
            this.TopMost = checkBoxAlwaysOnTop.Checked;
        }
    }
}
