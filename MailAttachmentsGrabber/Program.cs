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
            ILogger logger = new Logger();
            logger.LogToConsole = true;
            logger.LogFilePath = ".\\grab.log";

            MailGrabber mg = new MailGrabber();
            mg.Logger = logger;
            mg.UserName = "";
            mg.Password = "";
            mg.HostName = "https://outlook.office365.com/EWS/Exchange.asmx";
            mg.ConnectionSetup();

            // username, subjectprefix, attachmentnameprefix, numberofsearchmails
            mg.GetAttachments("bernt.anker@me.com", "", "IMG_", 100);

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
