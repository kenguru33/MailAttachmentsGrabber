using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailAttachmentsGrabber
{
    class Logger : ILogger
    {
        private TextWriter logFileWriter = null;
        private String logFilePath;
        private Boolean logtoconsole = true;


        public string LogFilePath
        {
            get
            {
                return this.logFilePath;
            }

            set
            {
                // TODO: Check if value is a valid file path.
                this.logFilePath = value;
            }
        }

        public bool LogToConsole
        {
            get
            {
                return this.logtoconsole;
            }

            set
            {
                this.logtoconsole = value;
            }
        }

        public void WriteLine(String logMessage)
        {
            if (this.logFilePath != null)
            {
                if (logFileWriter == null)
                {
                    this.logFileWriter = new StreamWriter(this.LogFilePath);
                }
                logFileWriter.WriteLine(DateTime.Now + " : " + logMessage);
            }
            Console.WriteLine(logMessage);
        }

        public void Clear()
        {
            // TODO: Write impl.
            throw new NotImplementedException();
        }
    }
}
