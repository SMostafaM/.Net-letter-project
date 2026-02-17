using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace daftar.Models
{
    public class Log
    {
        [BsonId]
        [BsonRepresentation(MongoDB.Bson.BsonType.ObjectId)]
        public string Id { get; set; }
        public string user { get; set; }
        public DateTime time { get; set; }
        public string action { get; set; }
        public string subject { get; set; }
        public string obj { get; set; }
    }
}
