using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.VisualBasic;
using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using Newtonsoft.Json.Linq;

namespace myPivotapp.Models
{
    public class PivotInputModel
    {
        [BsonId]
        public string Id { get; set; }
        
        
        [BsonExtraElements]
        public BsonDocument pivotInput { get; set; }
        //public string City { get; set; }



        //public JObject pivotInput { get; set; }
    }
}
