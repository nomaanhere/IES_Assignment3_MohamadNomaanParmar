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
    public static class FileSaver
    {
        public static void SaveUsersToExcel(List<User> users, string filePath)
        {
            FileInfo fileInfo = new FileInfo(filePath);
            ExcelPackage package = new ExcelPackage(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Users"];
            if (worksheet != null)
            {
                package.Workbook.Worksheets.Delete(worksheet);
            }
            worksheet = package.Workbook.Worksheets.Add("Users");

            int row = 1;
            worksheet.Cells[row, 1].Value = "Name";
            worksheet.Cells[row, 2].Value = "Email";
            worksheet.Cells[row, 3].Value = "Phone";
            worksheet.Cells[row, 4].Value = "Address";
            row++;

            foreach (var user in users)
            {
                worksheet.Cells[row, 1].Value = user.Name;
                worksheet.Cells[row, 2].Value = user.Email;
                worksheet.Cells[row, 3].Value = user.Phone;
                worksheet.Cells[row, 4].Value = $"{user.Address.Suite}, {user.Address.Street}, {user.Address.City}";
                row++;
            }

            package.Save();
        }
    }
}