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
    public static class DataDisplayer
    {
        public static void DisplayUsers(List<User> users)
        {
            Console.WriteLine($"{"Name",-25} {"Email",-25} {"Phone",-20} {"Address",-40}");
            Console.WriteLine(new string('=', 110));

            foreach (var user in users)
            {
                string address = $"{user.Address.Suite}, {user.Address.Street}, {user.Address.City}";
                Console.WriteLine($"{user.Name,-25} {user.Email,-25} {user.Phone,-20} {address,-40}");
            }
        }
    }
}