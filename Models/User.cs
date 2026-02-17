using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace daftar.Models
{
    public class User
    {
        [BsonId]
        [BsonRepresentation(MongoDB.Bson.BsonType.ObjectId)]
        public string Id { get; set; }

        [BsonElement("user_name")]
        [Display(Name = "نام کاربری")]
        public string user_name { get; set; }

        [BsonElement("password")]
        [DataType(DataType.Password)]
        [Display(Name = "کلمه عبور")]
        public string password { get; set; }
        public string role { get; set; }
        public string token { get; set; }
        
    }
}
