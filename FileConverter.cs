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
    public static class FileConverter
    {
        public static void CreateExcelFromXml(string xmlFilePath, string excelFilePath)
        {
            XDocument xmlDocument = XDocument.Load(xmlFilePath);
            var xmlBooks = xmlDocument.Descendants("book");

            FileInfo fileInfo = new FileInfo(excelFilePath);
            
            ExcelPackage package = new ExcelPackage(fileInfo);
            
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Books"];
            if (worksheet != null)
            {
                package.Workbook.Worksheets.Delete(worksheet);
            }
            worksheet = package.Workbook.Worksheets.Add("Books");

            int row = 1;
            worksheet.Cells[row, 1].Value = "Category";
            worksheet.Cells[row, 2].Value = "Title";
            worksheet.Cells[row, 3].Value = "Author";
            worksheet.Cells[row, 4].Value = "Year";
            worksheet.Cells[row, 5].Value = "Price";
            row++;

            foreach (var book in xmlBooks)
            {
                worksheet.Cells[row, 1].Value = book.Attribute("category")?.Value;
                worksheet.Cells[row, 2].Value = book.Element("title")?.Value;
                worksheet.Cells[row, 3].Value = book.Element("author")?.Value;
                worksheet.Cells[row, 4].Value = book.Element("year")?.Value;
                worksheet.Cells[row, 5].Value = book.Element("price")?.Value;
                row++;
            }

            package.Save();
        }

        public static void ConvertJsonToXml(string jsonFilePath, string xmlFilePath)
        {
            string jsonContent = File.ReadAllText(jsonFilePath);
            JObject json = JObject.Parse(jsonContent);
            XDocument xml = JsonConvert.DeserializeXNode(json.ToString(), "root");
            xml.Save(xmlFilePath);
        }

        public static void ConvertXmlToJson(string xmlFilePath, string jsonFilePath)
        {
            XDocument xml = XDocument.Load(xmlFilePath);
            string json = JsonConvert.SerializeXNode(xml);
            File.WriteAllText(jsonFilePath, json);
        }

        public static void ConvertCsvToExcel(string csvFilePath, string excelFilePath)
        {
            DataTable csvDataTable = new DataTable();
            StreamReader sr = new StreamReader(csvFilePath);
            string[] headers = sr.ReadLine().Split(',');
            foreach (string header in headers)
            {
                csvDataTable.Columns.Add(header);
            }
            while (!sr.EndOfStream)
            {
                string[] rows = sr.ReadLine().Split(',');
                DataRow dr = csvDataTable.NewRow();
                for (int i = 0; i < headers.Length; i++)
                {
                    dr[i] = rows[i];
                }
                csvDataTable.Rows.Add(dr);
            }

            FileInfo fileInfo = new FileInfo(excelFilePath);
            ExcelPackage package = new ExcelPackage(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
            if (worksheet != null)
            {
                package.Workbook.Worksheets.Delete(worksheet);
            }
            worksheet = package.Workbook.Worksheets.Add("Sheet1");

            for (int i = 0; i < csvDataTable.Columns.Count; i++)
            {
                worksheet.Cells[1, i + 1].Value = csvDataTable.Columns[i].ColumnName;
            }
            for (int i = 0; i < csvDataTable.Rows.Count; i++)
            {
                for (int j = 0; j < csvDataTable.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1].Value = csvDataTable.Rows[i][j];
                }
            }

            package.Save();
        }
    }
}