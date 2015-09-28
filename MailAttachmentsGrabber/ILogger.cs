using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailAttachmentsGrabber
{
    interface ILogger
    {
        String LogFilePath
        {
            get;
            set;
        }

        Boolean LogToConsole
        {
            get;
            set;
        }

        void WriteLine(String logText);
        void Clear();
    }
}
