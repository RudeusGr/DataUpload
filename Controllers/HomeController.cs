using ClosedXML.Excel;
using Dapper;
using DataUpload.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System.Diagnostics;

namespace DataUpload.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IConfiguration _configuration;

        public HomeController(ILogger<HomeController> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult LoadData(IFormFile fileExcel)
        {
            XLWorkbook workbook = new XLWorkbook(fileExcel.OpenReadStream());

            IXLWorksheet sheet = workbook.Worksheet(1);

            int firstrow = sheet.FirstRowUsed().RangeAddress.FirstAddress.RowNumber;
            int lastrow = sheet.LastRowUsed().RangeAddress.FirstAddress.RowNumber;

            List<Assistance> listAssistance = new List<Assistance>();

            for (int i = firstrow; i < lastrow; i++)
            {
                IXLRow row = sheet.Row(i);
                Assistance assistance = new Assistance();
                assistance.Date = row.Cell(1).GetValue<DateTime>();
                assistance.Route = row.Cell(2).GetValue<int>();
                assistance.CVEmploye = row.Cell(3).GetString();
                listAssistance.Add(assistance);
            }

            using (SqlConnection DBConnection = new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                string sqlQuery = "INSERT INTO assistances(date,route,cvemployee) VALUES (@Date,@Route,@CVEmploye)";
                var result = DBConnection.Execute(sqlQuery, listAssistance);
            }

            return View("Index");
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
