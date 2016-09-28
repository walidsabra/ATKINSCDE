using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Testing_App.classes
{
    class wfObject
    {
        public int WorkflowId { get; set; }

        public string WorkflowName { get; set; }
        public string StateName { get; set; }

        public int StateOneId { get; set; }

        public int StateTwoId { get; set; }

        public string SharedStateName { get; set; }
        public string ArchivedStateName {get; set;}

    }
}
