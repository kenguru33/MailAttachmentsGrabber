using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailAttachmentsGrabber
{
    class Program
    {
        static void Main(string[] args)
        {
            CommandLineParser cmdParser = new CommandLineParser(args);

            ILogger logger = new Logger();
            logger.LogToConsole = true;
            logger.LogFilePath = ".\\grab.log";

            MailGrabber mg = new MailGrabber();
            mg.Logger = logger;

            String output = "";
            if (cmdParser.Args.TryGetValue("host",out output))
                mg.HostName = output;
            
            if (cmdParser.Args.TryGetValue("username", out output))
                mg.UserName = output;
            
            if (cmdParser.Args.TryGetValue("password", out output))
                mg.Password = output;

            if (cmdParser.Args.TryGetValue("logfile", out output))
                mg.Logger.LogFilePath = output;

            if (cmdParser.Args.TryGetValue("verbose", out output))
                if (output == "yes")
                    mg.Logger.LogToConsole = true;
                else
                    mg.Logger.LogToConsole = false;

            String subjectPrefix;
            if (cmdParser.Args.TryGetValue("subjectprefix", out subjectPrefix) == false)
                subjectPrefix = ""; 

            String attachmentFilenamePrefix;
            if (cmdParser.Args.TryGetValue("attachmentFilenamePrefix", out attachmentFilenamePrefix) == false)
                attachmentFilenamePrefix = "";

            String mailCount;           
            if (cmdParser.Args.TryGetValue("count", out mailCount) == false)
                mailCount = "1000";

            int searchMailCount;
            if (int.TryParse(mailCount, out searchMailCount) == false)
                searchMailCount = 1000;
                
            mg.ConnectionSetup();

            // username, subjectprefix, attachmentnameprefix, numberofsearchmails
            mg.GetAttachments(mg.UserName, subjectPrefix, attachmentFilenamePrefix, searchMailCount);

            Console.WriteLine("Press ESC to stop");
            do
            {
                while (!Console.KeyAvailable)
                {
                    // Do something
                }
            } while (Console.ReadKey(true).Key != ConsoleKey.Escape);

        }
    }
}
