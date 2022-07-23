using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookCaptureEmailAddIn.Model.POCO
{
    public class Rule
    {
        public String Field { get; set; }
        public String Value { get; set; }
        public string Action { get; set; }
        public string Target { get; set; }
        public String StoreID { get; set; }
    }
}
