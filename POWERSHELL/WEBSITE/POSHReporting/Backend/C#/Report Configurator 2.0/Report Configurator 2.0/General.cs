using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PSReport
{
    public class General
    {
        public string FileName { get; set; }
        public string ReportName { get; set; }

        public General()
        {

        }

        public General(string filename, string reportname)
        {
            FileName = filename;
            ReportName = reportname;
        }

    }
}
