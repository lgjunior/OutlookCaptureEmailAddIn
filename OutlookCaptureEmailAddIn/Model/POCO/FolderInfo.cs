using MongoDB.Bson;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookCaptureEmailAddIn.Model.POCO
{
    public class FolderInfo
    {
        public ObjectId _id { get; set; }
        public String EntryID { get; set; }
        public String FolderPath { get; set; }
        public String StoreID { get; set; }
    }
}
