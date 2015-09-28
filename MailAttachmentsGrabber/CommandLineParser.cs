using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailAttachmentsGrabber
{
    class CommandLineParser
    {
        private Dictionary<String, String> args;

        public Dictionary<String, String> Args
        {
            get
            {
                return this.args;
            }
            set
            {
                this.args = value;
            }
        }

        public CommandLineParser(String[] args)
        {
            this.args = new Dictionary<string, string>();
            foreach (String arg in args)
            {
                string[] stringSeparator = new string[] { "=" };
                String[] s = arg.Split(new char[] { '=' } );
                this.args.Add(s[0], s[1]);
            }
        }
    }
}
