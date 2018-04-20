using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using ReadDataFromExcel.Models;
using Bytescout.Spreadsheet;
using System;
using System.Collections.Generic;

namespace ReadDataFromExcel.Controllers
{
    public class HomeController : Controller
    {
        
        public IActionResult Index()
        {
            Spreadsheet document = new Spreadsheet();
            document.LoadFromFile("C:\\temp\\test.xlsx");
            Worksheet worksheet = document.Workbook.Worksheets.ByName("Taul1");
            int STARTING_ROW = 2;

            List<UserModel> usersToBeAddedList = new List<UserModel>();
            // Check dates
            for (int i = STARTING_ROW; i < 10; i++)
           {
                
                // Set current cell
                Cell username = worksheet.Cell(i, 0);
                Cell password = worksheet.Cell(i, 1);
                Cell nimi = worksheet.Cell(i, 2);
                Cell sukunimi = worksheet.Cell(i, 3);
                Cell osoite = worksheet.Cell(i, 4);
                Cell email = worksheet.Cell(i, 5);
                Cell employeetype = worksheet.Cell(i, 6);
                Cell enabled = worksheet.Cell(i, 7);
                if (username.ValueAsString.Equals(""))
                    break;
                UserModel userModel = new UserModel
                {
                    Username = username.ValueAsString,
                    Nimi = nimi.ValueAsString,
                    Sukunimi = sukunimi.ValueAsString,
                    Password = password.ValueAsString,
                    Osoite = osoite.ValueAsString,
                    Email = email.ValueAsString,
                    EmployeeType = employeetype.ValueAsString,
                    Enabled = enabled.ValueAsBoolean
                };


                usersToBeAddedList.Add(userModel);
                // Write Date

            }

            // Close document
            document.Close();
            return View();
        }

        public IActionResult About()
        {
            ViewData["Message"] = "Your application description page.";

            return View();
        }

        public IActionResult Contact()
        {
            ViewData["Message"] = "Your contact page.";

            return View();
        }

        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
