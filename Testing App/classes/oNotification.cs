using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CDEAutomation.classes
{
    class oNotification
    {
        public string host { get; set; }
        public int port { get; set; }
        public bool ssl { get; set; }
        public string To { get; set; }
        public string From { get; set; }
        public string Subject { get; set; }
        public string Body { get; set;  }
        public string user {get; set;}
        public string password { get; set; }
    }
}
