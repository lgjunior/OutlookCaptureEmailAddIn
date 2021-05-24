using MongoDB.Bson;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookCaptureEmailAddIn.Model.POCO
{
    public class Delete
    {
        public ObjectId _id { get; set; }
        public string SenderEmailAddress { get; set; }
    }
}
