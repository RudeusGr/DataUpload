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
            var fileStrem = fileExcel.FileName.Split('.');
            if (!fileStrem[fileStrem.Length-1].Equals("xlsx"))
            {
                ViewBag.Error = "El archivo enviado no es un archivo de Excel valido, los archivos de Excel tiene la extencion 'xlsx'";
                return View("Index");
            }

            XLWorkbook workbook = new XLWorkbook(fileExcel.OpenReadStream());

            IXLWorksheet sheet = workbook.Worksheet(1);

            int firstrow = sheet.FirstRowUsed().RangeAddress.FirstAddress.RowNumber;
            int lastrow = sheet.LastRowUsed().RangeAddress.FirstAddress.RowNumber;

            using (SqlConnection DBConnection = new SqlConnection(_configuration.GetConnectionString("DefaultConnection")))
            {
                DBConnection.Open();
                string sqlQuery = "INSERT INTO assistances(date,route,cvemployee) VALUES (@Date,@Route,@CVEmploye)";

                using (var transaction = DBConnection.BeginTransaction())
                {
                    try
                    {
                        for (int i = firstrow; i < lastrow; i++)
                        {
                            IXLRow row = sheet.Row(i);
                            Assistance assistance = new Assistance();
                            assistance.Date = row.Cell(1).GetValue<DateTime>();
                            assistance.Route = row.Cell(2).GetValue<int>();
                            assistance.CVEmploye = row.Cell(3).GetString();
                            DBConnection.Execute(sqlQuery, assistance, transaction: transaction);
                        }
                        transaction.Commit();
                        ViewBag.Success = "Asistencias registradas correctamente.";
                    }
                    catch (Exception err)
                    {
                        ViewBag.Error = "Sucedio un Error al registrar las asistencias, por favor revise el archivo enviado. Error:" + err.Message;
                        transaction.Rollback();
                    }
                }

                DBConnection.Close();
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
