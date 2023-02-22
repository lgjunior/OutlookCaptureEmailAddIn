using MongoDB.Bson;
using MongoDB.Driver;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookCaptureEmailAddIn.Model
{
    public static class DB
    {
        public static MongoClient dbClient = new MongoClient(@"mongodb://root:Li1p4ss1!@localhost:27017/?authMechanism=SCRAM-SHA-1");
        public static IMongoDatabase db = dbClient.GetDatabase("outlook");

        // --- Create

        public static void CreateDelete(POCO.Delete _record)
        {
            ReadCollDelete().InsertOne(_record);
        }

        public static void CreateMove(POCO.Move _record)
        {
            ReadCollMove().InsertOne(_record);
        }

        // --- Read

        public static IMongoCollection<POCO.Delete> ReadCollDelete()
        {
            return db.GetCollection<POCO.Delete>("Delete");
        }

        public static IEnumerable<POCO.Delete> ReadAllDelete()
        {
            return ReadCollDelete().Find(new BsonDocument()).ToList();
        }

        public static IMongoCollection<POCO.Move> ReadCollMove()
        {
            return db.GetCollection<POCO.Move>("Move");
        }

        public static IEnumerable<POCO.Move> ReadAllMove()
        {
            return ReadCollMove().Find(new BsonDocument()).ToList();
        }

        // --- Update

        // --- Delete

        public static void DeleteColl(string _collection)
        {
            IMongoCollection<BsonDocument> coll = db.GetCollection<BsonDocument>(_collection);
            if (coll != null)
            {
                db.DropCollection(_collection);
            }
        }
    }
}
