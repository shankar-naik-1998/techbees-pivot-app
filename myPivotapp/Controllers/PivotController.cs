using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using myPivotapp.BussinesLogic;
using myPivotapp.Models;
using Newtonsoft.Json.Linq;
using System.IO;

using Newtonsoft.Json;
using Microsoft.VisualBasic;
using MongoDB.Bson;
using MongoDB.Bson.Serialization;
using System.Collections;
using System.Data;
using Microsoft.AspNetCore.Mvc.Diagnostics;

namespace myPivotapp.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class PivotController : ControllerBase
    {
        public  static IPivotServices _pivotServices;
       public  static JObject pivotInputJsonFIle;
        public PivotController(IPivotServices pivotServices)
        {
            _pivotServices = pivotServices;
            JObject pivotInputJsonFIle;
        }
       
        [HttpPost]
        public  IActionResult getPivot([FromForm] IFormFile Inputfile) {
            
            try
            {/*
                if (Inputfile == null || Inputfile.Length == 0)
                    return Content("file not selected");*/

                var filepath = Path.Combine(
                            Directory.GetCurrentDirectory(),
                            Inputfile.FileName);

                using (var stream = new FileStream(filepath, FileMode.Create))
                {
                    Inputfile.CopyToAsync(stream);
                }

                string path = filepath;

                //Use when you can save the file-- string extension = Path.GetExtension(Inputfile.FileName);
               // string path = @"C:\pivotInputFiles\SalesJan2009.xlsx";
                string extension = Path.GetExtension(path);

                //string inputJsonString = string.Empty;
                DataTable dt = new DataTable();

                if(extension.ToString().ToUpper()=="CSV")
                {

                }
                else if(extension.ToString().ToUpper()==".XLSX")
                {
                    dt = _pivotServices.ExcelToDataTable(path);
                    
                }
                //inputJsonString = JsonConvert.SerializeObject(pivotInputJsonFIle);
                //string text = System.IO.File.ReadAllText(@"C:\Users\user\pivotInputData.json");
               


                /*IEnumerable<BsonDocument> bsonIEnumerable = new List<BsonDocument>();
                string[] separatingChars = { "[","]" }; // split on these chars
                string[] docs = inputJsonString.Split(separatingChars, System.StringSplitOptions.RemoveEmptyEntries);
                foreach(var doc in docs)
                {*/
                    //var file = BsonSerializer.Deserialize<BsonDocument>(inputJsonString);
                  
                   

               // }
                //List<BsonDocument> result=_pivotServices.Create(dt);

                return Ok();



                /*var  input = System.IO.File.ReadAllText(@"C:\Users\user\pivotInputData.json");
                JObject result = JObject.Parse(input);*/



                /*PivotInputModel data = new PivotInputModel
                {
                    Id = Guid.NewGuid().ToString(),
                    pivotInput = BsonDocument.Parse(result.ToString())

                };*/

            }
            catch(Exception e)
            {
                return BadRequest("Some thing went wrong Please try Agian..!");

            }
    }

    }
}
