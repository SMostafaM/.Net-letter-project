using MongoDB.Driver;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace daftar
{
    public class Context
    {
        private readonly IMongoDatabase database;
        public Context()
        {
            database = new MongoClient("mongodb://localhost:27017").GetDatabase("daftar");
        }
        
        public IMongoCollection<Models.User> User
        {
            get
            {
                return database.GetCollection<Models.User>("User");
            }

        }
        public IMongoCollection<Models.Log> Log
        {
            get
            {
                return database.GetCollection<Models.Log>("Log");
            }

        }
        public IMongoCollection<Models.Letter> Letter
        {
            get
            {
                return database.GetCollection<Models.Letter>("Letter");
            }

        }
        public IMongoCollection<Models.Letter> deleted_Letter
        {
            get
            {
                return database.GetCollection<Models.Letter>("deleted_Letter");
            }

        }
        public IMongoCollection<Models.Sanction> sanction
        {
            get
            {
                return database.GetCollection<Models.Sanction>("sanction");
            }

        }
        public IMongoCollection<Models.Sanction> deleted_sanction
        {
            get
            {
                return database.GetCollection<Models.Sanction>("deleted_sanction");
            }

        }
        public IMongoCollection<Models.Active_session> Active_session
        {
            get
            {
                return database.GetCollection<Models.Active_session>("Active_session");
            }

        }

    }
}
