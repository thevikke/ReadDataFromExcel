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
           while(true)
           {
                
                // Set current cell
                Cell username = worksheet.Cell(STARTING_ROW, 0);
                Cell password = worksheet.Cell(STARTING_ROW, 1);
                Cell nimi = worksheet.Cell(STARTING_ROW, 2);
                Cell sukunimi = worksheet.Cell(STARTING_ROW, 3);
                Cell osoite = worksheet.Cell(STARTING_ROW, 4);
                Cell email = worksheet.Cell(STARTING_ROW, 5);
                Cell employeetype = worksheet.Cell(STARTING_ROW, 6);
                Cell enabled = worksheet.Cell(STARTING_ROW, 7);
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
                STARTING_ROW += 1;
            }

            // Close document
            document.Close();
            return View(usersToBeAddedList);
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
