using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Xml.Linq;

namespace JSONXMLFunc
{
    public static class DataFetcher
    {
        public static async Task<List<User>> GetUserDataAsync(string url)
        {
            HttpClient client = new HttpClient();
            var response = await client.GetStringAsync(url);
            return JsonConvert.DeserializeObject<List<User>>(response);
        }
    }
}