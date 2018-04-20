using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ReadDataFromExcel.Models
{
    public class UserModel
    {
        public string Username{ get; set; }
        public string Password { get; set; }
        public string Nimi { get; set; }
        public string Sukunimi{ get; set; }
        public string Osoite { get; set; }
        public string Email { get; set; }
        public string EmployeeType { get; set; }
        public bool Enabled { get; set; }
    }
}
