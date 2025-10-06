using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemoAddInTC.model
{
    class BOMLine
    {
        public String FullName { get; set; }
        public String level { get; set; }
        public String AbsolutePath { get; set; }
        public String DocNum { get; set; }
        public String Revision { get; set; }
        public String Status { get; set; }
    }
}
