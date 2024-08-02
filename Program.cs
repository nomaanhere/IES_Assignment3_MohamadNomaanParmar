using System;
using System.Threading.Tasks;
using OfficeOpenXml;
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
    class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Set EPPlus license context

            while (true)
            {
                Console.WriteLine("Select a task to perform:");
                Console.WriteLine("1. Fetch and display user data");
                Console.WriteLine("2. Save user data to Excel");
                Console.WriteLine("3. Display Books data from JSON, CSV, and XML");
                Console.WriteLine("4. Generate Excel file from XML");
                Console.WriteLine("5. Convert JSON to XML");
                Console.WriteLine("6. Convert XML to JSON");
                Console.WriteLine("7. Convert CSV to Excel");
                Console.WriteLine("8. Convert XML to Excel");
                Console.WriteLine("9. Exit");
                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        var users = await DataFetcher.GetUserDataAsync("https://jsonplaceholder.typicode.com/users");
                        DataDisplayer.DisplayUsers(users);
                        break;
                    case "2":
                        users = await DataFetcher.GetUserDataAsync("https://jsonplaceholder.typicode.com/users");
                        FileSaver.SaveUsersToExcel(users, FileHelper.GetDesktopPath("users.xlsx"));
                        Console.WriteLine("User data saved to users.xlsx");
                        break;
                    //case "3":
                      //  DataReader.DisplayBooksData();
                        //break;
                    //case "4":
                        FileConverter.CreateExcelFromXml(FileHelper.GetDesktopPath("Books.xml"), FileHelper.GetDesktopPath("books.xlsx"));
                        Console.WriteLine("Books data saved to books.xlsx");
                        break;
                    case "5":
                        FileConverter.ConvertJsonToXml(FileHelper.GetDesktopPath("Books.json"), FileHelper.GetDesktopPath("BooksConverted.xml"));
                        Console.WriteLine("JSON data converted to XML and saved as BooksConverted.xml");
                        break;
                    case "6":
                        FileConverter.ConvertXmlToJson(FileHelper.GetDesktopPath("Books.xml"), FileHelper.GetDesktopPath("BooksConverted.json"));
                        Console.WriteLine("XML data converted to JSON and saved as BooksConverted.json");
                        break;
                    case "7":
                        FileConverter.ConvertCsvToExcel(FileHelper.GetDesktopPath("Books.csv"), FileHelper.GetDesktopPath("BooksConverted.xlsx"));
                        Console.WriteLine("CSV data converted to Excel and saved as BooksConverted.xlsx");
                        break;
                    case "8":
                        FileConverter.CreateExcelFromXml(FileHelper.GetDesktopPath("Books.xml"), FileHelper.GetDesktopPath("BooksConverted.xlsx"));
                        Console.WriteLine("XML data saved to BooksConverted.xlsx");
                        break;
                    case "9":
                        return;
                    default:
                        Console.WriteLine("Invalid choice. Please select a valid option.");
                        break;
                }
                Console.WriteLine("Press any key to return to the main menu...");
                Console.ReadKey();
                Console.Clear();
            }
        }
    }
}