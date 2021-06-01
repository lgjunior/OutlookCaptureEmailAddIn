using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookCaptureEmailAddIn.Model.POCO
{
    public class Rule
    {
        public string Email { get; set; }
        public string Action { get; set; }
        public string Target { get; set; }
        public String StoreID { get; set; }
    }
}
