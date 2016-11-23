using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Serilog;
using InfluxDB.Collector;

namespace ForwardMailsLibrary
{
    public class ForwardMails
    {

        private string mailboxSMTP;
        private ExchangeService service;
        private Folder srcFolder;
        private Folder dstFolder;


        public ForwardMails(string mailboxSMTP, string srcFolderName, string dstFolderName, bool impersonate, ExchangeVersion requestedServerVersion = ExchangeVersion.Exchange2010_SP2)
        {
            try
            {
                Log.Debug("ExchangeService: Creating - mailboxSMTP: {mailboxSMTP}  - srcFolderName: {srcFolderName} -  dstFolderName: {dstFolderName}", mailboxSMTP, srcFolderName, dstFolderName);
                ExchangeService service = new ExchangeService(requestedServerVersion);
                service.UseDefaultCredentials = true;
                service.AutodiscoverUrl(mailboxSMTP);
                if(impersonate)
                    service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, mailboxSMTP);
                this.service = service;
                this.mailboxSMTP = mailboxSMTP;
                Log.Debug("ExchangeService: Ready");

            }
            catch (Exception ex)
            {
                Log.Error(ex, "Failed while trying to create ExchangeService");
                throw;
            }

            try
            {
                Mailbox mb = new Mailbox(mailboxSMTP);
                FolderId fIdMsgRoot = new FolderId(WellKnownFolderName.Inbox, mb);

                Folder rootfolder = Folder.Bind(service, fIdMsgRoot);
                Log.Debug("The {rootfolderDisplayName} has {rootfolderChildFolderCount} child folders.", rootfolder.DisplayName, rootfolder.ChildFolderCount);

                PropertySet propSet = new PropertySet(BasePropertySet.IdOnly, FolderSchema.DisplayName);
                rootfolder.Load(propSet);

                foreach (Folder folder in rootfolder.FindFolders(new FolderView(100)))
                {
                    Log.Debug("Name: {folderDisplayName}. Id: {folderId}", folder.DisplayName, folder.Id);                    
                    
                    if (folder.DisplayName == srcFolderName)
                    {
                        Log.Debug("This is the source folder: {folderDisplayName}", folder.DisplayName);
                        srcFolder = folder;
                    }

                    if (folder.DisplayName == dstFolderName)
                    {
                        Log.Debug("This is the destination folder: {folderDisplayName}", folder.DisplayName);
                        dstFolder = folder;
                    }
                }

                if (srcFolder == null)
                {
                    throw new Exception(string.Format("Folder {0} not found", srcFolderName));
                }
                if (dstFolder == null)
                {
                    throw new Exception(string.Format("Folder {0} not found", dstFolderName));
                }

            }
            catch (Exception ex)
            {
                Log.Error(ex, "Failed to select folder");
                throw ex;
            }

            
        }


        public void ForwardMailsInFolder(string forwardAddress, bool delete)
        {
            Log.Debug("Entering ForwardMailInFolder with delete = {delete}", delete);
            Metrics.Increment("mainForwardLoopIteration");

            int offset = 0;
            int pageSize = 50;
            bool more = true;

            ItemView view = new ItemView(pageSize, offset, OffsetBasePoint.Beginning);
            view.PropertySet = PropertySet.IdOnly;

            FindItemsResults<Item> findResults;
            List<ItemId> allResultsID = new List<ItemId>();
            List<EmailMessage> emails = new List<EmailMessage>();

            try
            {
                Log.Debug("Start mails query");
                while (more)
                {
                    findResults = service.FindItems(srcFolder.Id, view);
                    Log.Debug("Pass {FindItemIteration} - findResults: {findResults}", view.Offset / pageSize, findResults.Count());
                    foreach (var item in findResults.Items)
                    {
                        emails.Add((EmailMessage)item);
                    }
                    more = findResults.MoreAvailable;
                    if (more)
                    {
                        view.Offset += pageSize;
                    }
                }

                if (emails.Count > 0)
                {
                    Log.Debug("{nbEmails} founds. Get properties.", emails.Count);
                    
                    Metrics.Increment("mailsFoundLoopIteration");
                    Metrics.Measure("nbMailsToForward", emails.Count);
                    
                    PropertySet properties = new PropertySet(BasePropertySet.FirstClassProperties);
                    service.LoadPropertiesForItems(emails, properties);

                    Log.Debug("Forward enabled: {forward}", true);

                    foreach (EmailMessage msg in emails)
                    {
                        Log.Information("Message with subject: {Subject} will be forwarded to {forwardAddress}", msg.Subject, forwardAddress);
                        allResultsID.Add(msg.Id);
                        msg.Body.BodyType = BodyType.HTML;
                        string messageBodyPrefix = GenerateMessageBodyPrefix(msg);
                        EmailAddress[] addresses = new EmailAddress[1];
                        addresses[0] = new EmailAddress(forwardAddress);

                        try
                        {
                            Log.Debug("Forward mail to: {forwardAddress}", forwardAddress);
                            Metrics.Increment("mailForwarded");
                            msg.Forward(messageBodyPrefix, addresses);
                        }
                        catch (Exception ex)
                        {
                            Log.Error(ex, "Failed to forward mails.");
                            Metrics.Increment("mailForwardError");
                            throw;
                        }
                        try
                        {
                            Log.Debug("Move mail to: {dstFolder}", dstFolder.DisplayName);
                            msg.Move(dstFolder.Id);
                        }
                        catch (Exception ex)
                        {
                            Log.Error(ex, "Failed to move mails.");
                            Metrics.Increment("mailMoveError");
                            throw;
                        }                    
                        
                    }

                    Log.Debug("Delete mail? {DeleteMail}", delete);
                    if (delete)
                    {
                        int nbMailsToDelete = allResultsID.Count;
                        Log.Information("Start Bulk Delete of {nbMailsToDelete} mails", nbMailsToDelete);
                        try
                        {
                            service.DeleteItems(allResultsID, DeleteMode.SoftDelete, SendCancellationsMode.SendToNone, AffectedTaskOccurrence.AllOccurrences);
                            Metrics.Measure("MailDeleted", nbMailsToDelete);
                        }
                        catch (Exception ex)
                        {

                            Log.Error(ex, "Failed to delete mails.");
                            Metrics.Increment("mailDeleteError");
                            throw;
                        }
                        
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Failed to forward, move or delete mails.");
                Metrics.Increment("mailProcessError");
                throw;
            }

        }

        private string GenerateMessageBodyPrefix(EmailMessage message)
        {
            string strFrom = message.From.Address;
            string strTo = string.Empty;
            string strCC = string.Empty;

            foreach (var toAddress in message.ToRecipients)
            {
                strTo += string.Format("{0};",toAddress.Address);
            }

            foreach (var ccAddress in message.CcRecipients)
            {
                strCC += string.Format("{0};", ccAddress.Address);
            }

            string msgBodyPrefix = string.Format("From:[mailto:{0}]",strFrom);

            if (strTo != string.Empty)
            {
                strTo = string.Format("To:{0}", strTo.Substring(0, strTo.Length - 1));
                msgBodyPrefix = string.Format("{0}{1}{2}", msgBodyPrefix, "<br>", strTo);
            }
            if (strCC != string.Empty)
            {
                strCC = string.Format("Cc:{0}", strCC.Substring(0, strCC.Length - 1));
                msgBodyPrefix = string.Format("{0}{1}{2}", msgBodyPrefix, "<br>", strCC);
            }

            msgBodyPrefix = string.Format("{0}{1}FW:{2}", msgBodyPrefix, "<br>", message.Subject);

            return msgBodyPrefix;
        }

    }
}
