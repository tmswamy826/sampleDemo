using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using Newtonsoft.Json.Linq;

namespace TestConsole
{
    class convertExceldataToJsondata
    {
        public void convertdata()
        {
            try
            {


                var pathToExcel = @"C:\Users\52045046\Desktop\Test.xlsx";
                var sheetName = "sheetOne";

                //This connection string works if you have Office 2007+ installed and your 
                //data is saved in a .xlsx file
                var connectionString = String.Format(@"
            Provider=Microsoft.ACE.OLEDB.12.0;
            Data Source={0};
            Extended Properties=""Excel 12.0 Xml;HDR=YES""
        ", pathToExcel);

                //Creating and opening a data connection to the Excel sheet 
                using (var conn = new OleDbConnection(connectionString))
                {
                    conn.Open();

                    var cmd = conn.CreateCommand();
                    cmd.CommandText = String.Format(
                        @"SELECT * FROM [Sheet1$]",
                        sheetName
                        );


                    using (var rdr = cmd.ExecuteReader())
                    {
                        //LINQ query - when executed will create anonymous objects for each row
                        var query =
                            (from DbDataRecord row in rdr
                             select row).Select(x =>
                             {


                                 //dynamic item = new ExpandoObject();
                                 Dictionary<string, object> item = new Dictionary<string, object>();
                                 item.Add(rdr.GetName(0), x[0]);
                                 item.Add(rdr.GetName(1), x[1]);
                                 //item.Add(rdr.GetName(2), x[2]);
                                 return item;

                             });

                        //Generates JSON from the LINQ query
                        var json = JsonConvert.SerializeObject(query);
                        // return json;
                        Console.WriteLine("Intents : " +json.ToString());
                    }

                }
            }

            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }

        }

        
        enum SplitState
        {
            InPrefix,
            InSplitProperty,
            InSplitArray,
            InPostfix,
        }

        public static void SplitJson(TextReader textReader, string tokenName, long maxItems, Func<int, TextWriter> createStream, Formatting formatting)
        {
            List<JProperty> prefixProperties = new List<JProperty>();
            List<JProperty> postFixProperties = new List<JProperty>();
            List<JsonWriter> writers = new List<JsonWriter>();

            SplitState state = SplitState.InPrefix;
            long count = 0;

            try
            {
                using (var reader = new JsonTextReader(textReader))
                {
                    bool doRead = true;
                    while (doRead ? reader.Read() : true)
                    {
                        doRead = true;
                        if (reader.TokenType == JsonToken.Comment || reader.TokenType == JsonToken.None)
                            continue;
                        if (reader.Depth == 0)
                        {
                            if (reader.TokenType != JsonToken.StartObject && reader.TokenType != JsonToken.EndObject)
                                throw new JsonException("JSON root container is not an Object");
                        }
                        else if (reader.Depth == 1 && reader.TokenType == JsonToken.PropertyName)
                        {
                            if ((string)reader.Value == tokenName)
                            {
                                state = SplitState.InSplitProperty;
                            }
                            else
                            {
                                if (state == SplitState.InSplitProperty)
                                    state = SplitState.InPostfix;
                                var property = JProperty.Load(reader);
                                doRead = false; // JProperty.Load() will have already advanced the reader.
                                if (state == SplitState.InPrefix)
                                {
                                    prefixProperties.Add(property);
                                }
                                else
                                {
                                    postFixProperties.Add(property);
                                }
                            }
                        }
                        else if (reader.Depth == 1 && reader.TokenType == JsonToken.StartArray && state == SplitState.InSplitProperty)
                        {
                            state = SplitState.InSplitArray;
                        }
                        else if (reader.Depth == 1 && reader.TokenType == JsonToken.EndArray && state == SplitState.InSplitArray)
                        {
                            state = SplitState.InSplitProperty;
                        }
                        else if (state == SplitState.InSplitArray && reader.Depth == 2)
                        {
                            if (count % maxItems == 0)
                            {
                                var writer = new JsonTextWriter(createStream(writers.Count)) { Formatting = formatting };
                                writers.Add(writer);
                                writer.WriteStartObject();
                                foreach (var property in prefixProperties)
                                    property.WriteTo(writer);
                                writer.WritePropertyName(tokenName);
                                writer.WriteStartArray();
                            }
                            count++;
                            writers.Last().WriteToken(reader, true);
                        }
                        else
                        {
                            throw new JsonException("Internal error");
                        }
                    }
                }
                foreach (var writer in writers)
                    using (writer)
                    {
                        writer.WriteEndArray();
                        foreach (var property in postFixProperties)
                            property.WriteTo(writer);
                        writer.WriteEndObject();
                    }
            }
            finally
            {
                // Make sure files are closed in the event of an exception.
                foreach (var writer in writers)
                    using (writer)
                    {
                    }

            }
        }


        public static void TestSplitJson(string json, string tokenName)
        {
            var builders = new List<StringBuilder>();
            using (var reader = new StringReader(json))
            {
                SplitJson(reader, tokenName, 2, i => { builders.Add(new StringBuilder()); return new StringWriter(builders.Last()); }, Formatting.Indented);
            }
            foreach (var s in builders.Select(b => b.ToString()))
            {
                Console.WriteLine(s);
            }
        }


        public void LoadJson()
        {
            using (StreamReader r = new StreamReader(@"D:\Mani\TestConsole\TestConsole\file.json"))
            {
                string json = r.ReadToEnd();
                //List<Item> items = JsonConvert.DeserializeObject<List<Item>>(json);
                TestSplitJson(json, "Travel Agent");
            }
        }

    }
    
}
