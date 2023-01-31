using DataStatistik.Data;
using Microsoft.AspNetCore.Mvc;
using ClosedXML.Excel;
using DataStatistik.Models;
using IHostingEnvironment = Microsoft.AspNetCore.Hosting.IHostingEnvironment;
using MimeKit;
using SqlConnection = Microsoft.Data.SqlClient.SqlConnection;
using SqlCommand = Microsoft.Data.SqlClient.SqlCommand;
using System.Data;
using System.Net.Mail;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DataStatistik.Controllers
{
    public class GenerateController : Controller
    {
        public IConfiguration _Configuration { get; set; }
        private readonly DataContext _context;
        private IHostingEnvironment _IHosting;
        public GenerateController(DataContext context, IHostingEnvironment iHosting, IConfiguration configuration)
        {
            _context = context;
            _IHosting = iHosting;
            _Configuration = configuration;
        }
        public IActionResult Index()
        {
            return View(_context.data_statistik_saved.ToList());
        }
        public IActionResult Generate(DateTime from, DateTime until, string btnsubmit)
        {
            var data = from d in _context.data_statistik_saved select d;
            ViewData["CurrentFilter"] = from;
            ViewData["CurrentFilters"] = until;

            if (btnsubmit == "Search")
            {
                if (from != null && until != null)
                {
                    data = data.Where(d => d.Date >= from && d.Date <= until);
                }
            }
            if(btnsubmit == "Generate")
            {
                List<data_statistik_saved> ds = _context.data_statistik_saved.Select(x => new data_statistik_saved
                {
                    Date = x.Date,
                    Period = x.Period,
                    MemberCode = x.MemberCode,
                    MemberName = x.MemberName,
                    Province = x.Province,
                    City = x.City,
                    Frequency = x.Frequency,
                    Volume = x.Volume,
                    Value = x.Value
                }).ToList();
                IEnumerable<data_statistik_saved> today = from data_statistik_saved in ds where 
                                                            data_statistik_saved.Date == DateTime.Today select data_statistik_saved;
                

                using var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Data Statistik");
                var currentRow = 1;

                worksheet.Cell(currentRow, 1).Value = "TGL";
                worksheet.Cell(currentRow, 2).Value = "Kode Member";
                worksheet.Cell(currentRow, 3).Value = "Nama Member";
                worksheet.Cell(currentRow, 4).Value = "Domisili Provinsi";
                worksheet.Cell(currentRow, 5).Value = "Domisili Kota";
                worksheet.Cell(currentRow, 6).Value = "Frekuensi";
                worksheet.Cell(currentRow, 7).Value = "Volume";
                worksheet.Cell(currentRow, 8).Value = "Value";

               

                foreach (var user in today)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = user.Date.ToString("yyyy-MM-dd");
                    worksheet.Cell(currentRow, 2).Value = user.MemberCode;
                    worksheet.Cell(currentRow, 3).Value = user.MemberName;
                    worksheet.Cell(currentRow, 4).Value = user.Province;
                    worksheet.Cell(currentRow, 5).Value = user.City;
                    worksheet.Cell(currentRow, 6).Value = user.Frequency;
                    worksheet.Cell(currentRow, 7).Value = user.Volume;
                    worksheet.Cell(currentRow, 8).Value = user.Value;


                }
                var location = _IHosting.WebRootPath + "/Text/";
                var filename = location + $"Data Trading Bursa Harian {DateTime.Now.ToString("yyyyMMdd")}.xlsx";
                using var stream = new MemoryStream();
                var content = stream.ToArray();
                workbook.SaveAs(filename);
            }
            return View("Index", data);
        }
    }
}
