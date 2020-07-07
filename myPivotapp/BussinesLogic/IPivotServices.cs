using Microsoft.AspNetCore.Http;
using MongoDB.Bson;
using myPivotapp.Models;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq.Expressions;
using System.Threading.Tasks;

namespace myPivotapp.BussinesLogic
{
    public interface IPivotServices
    {
        List<BsonDocument> Create(DataTable pivotInput);
        DataTable ExcelToDataTable(string filepath);

      
    }
}