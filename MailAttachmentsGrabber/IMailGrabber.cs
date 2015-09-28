using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;

namespace MailAttachmentsGrabber
{
    interface IMailGrabber
    {
        // Gets and sets the logger file.
        ILogger Logger
        {
            get;
            set;
        }

        String HostName
        {
            get;
            set;
        }

        String UserName
        {
            get;
            set;
        }

        String Password
        {
            set;
        }

        String ExportPath
        {
            get;
            set;
        }

        // Get all attachements.
        void GetAttachments(String fromAddress, String subjectPrefix, string attachmentNamePrefix, int mailSearchCount);

        void ConnectionSetup();
          
    }
}
