using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;

namespace MailAttachmentsGrabber
{
    class MailGrabber : IMailGrabber
    {
        private String hostname;
        private String username;
        private String password;
        private String exportpath = ".\\";
        private ExchangeService service;
        private ILogger logger;

        public ILogger Logger
        {
            get
            {
                return this.logger;
            }

            set
            {
                this.logger = value; 
            }
        }

        public string HostName
        {
            get
            {
                return this.hostname;
            }

            set
            {
                this.hostname = value;
            }
        }

        public string UserName
        {
            get
            {
                return this.username;
            }

            set
            {
                this.username = value;
            }
        }

        public string Password
        {
            set
            {
                this.password = value;
            }
        }

        public string ExportPath
        {
            get
            {
                return this.exportpath;
            }

            set
            {
                this.exportpath = value;
            }
        }

        public void GetAttachments(String fromAddress, String subjectPrefix, String attachementNamePrefix, int mailSearchCount)
        {
            this.log("Searching inbox...");

            List<SearchFilter> searchFilterCollection = new List<SearchFilter>();
            searchFilterCollection.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.From, new EmailAddress(fromAddress)));
            if (subjectPrefix != "")
            {
                searchFilterCollection.Add(new SearchFilter.ContainsSubstring(EmailMessageSchema.Subject, subjectPrefix, ContainmentMode.Prefixed, ComparisonMode.IgnoreCase));
            }
            SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, searchFilterCollection.ToArray());

            FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, searchFilter, new ItemView(mailSearchCount));

            this.log("Emails found mathcing search criteria (From=" + fromAddress + " and subject prefix=" + subjectPrefix + ")=" + findResults.TotalCount.ToString());
            this.log("Searching for attachments...");

            foreach (Item item in findResults.Items)
            {
                if (item.HasAttachments)
                {                    
                    EmailMessage message = EmailMessage.Bind(this.service, item.Id, new PropertySet(BasePropertySet.IdOnly, ItemSchema.Attachments));
                    this.log(message.Attachments.Count() + " attachements found in email with ID=" + item.Id);
                    message.Load();
                    this.log("Searching for attachement matching file name prefix=" + attachementNamePrefix + "...");
                    Boolean filesFound = false;
                    foreach (FileAttachment fileAttachment in message.Attachments)
                    {
                        this.log("- - - - - - - - - - - - - Attachment - - - - - - - - - - - - - -");
                        fileAttachment.Load();
                        if (fileAttachment.Name.StartsWith(attachementNamePrefix))
                        {
                            filesFound = true;
                            this.log("Match found on fileName: " + fileAttachment.Name);
                            String path = this.exportpath + fileAttachment.Name;
                            String temppath = path;

                            // rename file by index if exist.
                            int filecount = 0;
                            while (System.IO.File.Exists(temppath))
                            {
                                filecount++;
                                this.log("A file with this name already exist. Trying adding count index " + filecount + " to filename.");
                                temppath = System.IO.Path.GetFileNameWithoutExtension(path) + "_" + filecount.ToString() + System.IO.Path.GetExtension(path);
                            }
                            path = temppath;
                            
                            fileAttachment.Load(path);
                            this.log("Exported attachent to: " + path);
                        }
                    }
                    if (!filesFound)
                    {
                        this.log("No attachment files found!");
                    }
                    this.log("- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ");
                }
            }
        }

        public void ConnectionSetup()
        {
            this.log("Exchange server connnection setup started...");
            // Create the binding.
            this.service = new ExchangeService();

            // Set the credentials for the on-premises server.
            service.Credentials = new WebCredentials(this.username, this.password);

            this.log("Username set to: " + this.UserName);

            // Set the URL.

            if (this.HostName != null)
            {
                service.Url = new Uri(this.HostName);
                this.log("Exhange Server set to: " + this.HostName);
            }
            else
            {
                this.log("Exchange Server URL not set.");
                try
                {
                    this.log("Running auto discover...");
                    service.AutodiscoverUrl(this.username, this.RedirectionCallback);
                    this.log("Exchange server " + service.Url + " discovered.");
                }
                catch (Exception e)
                {
                    this.logger.WriteLine("Discovery error: " + e.Message);
                    throw e;
                }
            }
         }

        private void log(String logMessage)
        {
            if (this.logger != null)
            {
                this.logger.WriteLine(logMessage);
            }
        }

        private bool RedirectionCallback(string url)
        {
            // Return true if the URL is an HTTPS URL.
            if (url.ToLower().StartsWith("https://"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
