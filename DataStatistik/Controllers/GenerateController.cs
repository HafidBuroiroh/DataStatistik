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
using DocumentFormat.OpenXml.Bibliography;

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

                var currentMonth = DateTime.Now.Month;
                IEnumerable<data_statistik_saved> month = from data_statistik_saved in ds
                                                          where
                                                            data_statistik_saved.Date.Month == currentMonth
                                                          select data_statistik_saved;
                using var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Data Statistik Harian");
                var currentRow = 1;

                worksheet.Cell(currentRow, 1).Value = "Tanggal";
                worksheet.Cell(currentRow, 2).Value = "Kode Member";
                worksheet.Cell(currentRow, 3).Value = "Nama Member";
                worksheet.Cell(currentRow, 4).Value = "Domisili Provinsi";
                worksheet.Cell(currentRow, 5).Value = "Domisili Kota";
                worksheet.Cell(currentRow, 6).Value = "Frekuensi";
                worksheet.Cell(currentRow, 7).Value = "Volume";
                worksheet.Cell(currentRow, 8).Value = "Value";
                worksheet.Column(1).Width = 10;
                worksheet.Column(2).Width = 12;
                worksheet.Column(3).Width = 15;
                worksheet.Column(4).Width = 15;
                worksheet.Column(5).Width = 15;
                worksheet.Column(6).Width = 10;
                worksheet.Column(7).Width = 10;
                worksheet.Column(8).Width = 10;
                foreach (var datas in today)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = datas.Date.ToString("yyyyMMMMdd");
                    worksheet.Cell(currentRow, 2).Value = datas.MemberCode;
                    worksheet.Cell(currentRow, 3).Value = datas.MemberName;
                    worksheet.Cell(currentRow, 4).Value = datas.Province;
                    worksheet.Cell(currentRow, 5).Value = datas.City;
                    worksheet.Cell(currentRow, 6).Value = datas.Frequency;
                    worksheet.Cell(currentRow, 7).Value = datas.Volume;
                    worksheet.Cell(currentRow, 8).Value = datas.Value;


                }
                var location = _IHosting.WebRootPath + "/Text/";
                var filename = location + $"Data Trading Bursa harian{DateTime.Now.ToString("ddMMyyyy")}.xlsx";
                using var stream = new MemoryStream();
                var content = stream.ToArray();
                workbook.SaveAs(filename);


                using var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("Data Statistik Bulanan");
                var cR = 1;

                ws.Cell(cR, 1).Value = "Tanggal";
                ws.Cell(cR, 2).Value = "Kode Member";
                ws.Cell(cR, 3).Value = "Nama Member";
                ws.Cell(cR, 4).Value = "Domisili Provinsi";
                ws.Cell(cR, 5).Value = "Domisili Kota";
                ws.Cell(cR, 6).Value = "Frekuensi";
                ws.Cell(cR, 7).Value = "Volume";
                ws.Cell(cR, 8).Value = "Value";

                ws.Column(1).Width = 10;
                ws.Column(2).Width = 12;
                ws.Column(3).Width = 15;
                ws.Column(4).Width = 15;
                ws.Column(5).Width = 15;
                ws.Column(6).Width = 10;
                ws.Column(7).Width = 10;
                ws.Column(8).Width = 10;

                foreach (var user in month)
                {
                    cR++;
                    ws.Cell(cR, 1).Value = user.Date.ToString("yyyyMMMMdd");
                    ws.Cell(cR, 2).Value = user.MemberCode;
                    ws.Cell(cR, 3).Value = user.MemberName;
                    ws.Cell(cR, 4).Value = user.Province;
                    ws.Cell(cR, 5).Value = user.City;
                    ws.Cell(cR, 6).Value = user.Frequency;
                    ws.Cell(cR, 7).Value = user.Volume;
                    ws.Cell(cR, 8).Value = user.Value;


                }
                var namafile = location + $"Data Trading Bursa Bulanan{DateTime.Now.ToString("ddMMyyyy")}.xlsx";
                using var ss = new MemoryStream();
                var con = stream.ToArray();
                wb.SaveAs(namafile);

                var date = DateTime.Now;
                var body = new BodyBuilder();
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress("ujangng09@gmail.com");
                    mail.To.Add("hafidburoiroh09@gmail.com");
                    mail.Subject = $"Status Generate Excel Data Demografi Investor Data Trading Bursa_{date.ToString("yyyy-MM")}";
                    mail.Body = $"Dengan Hormat,\r\n\r\nDengan ini kami informasikan bahwa per tanggal {date.ToString("dd-MM-yyyy")} proses generate Excel Data Trading Bursa dinyatakan sukses";
                    DirectoryInfo dir = new DirectoryInfo(location);
                    foreach (FileInfo file in dir.GetFiles("*.*"))
                    {
                        if (file.Exists)
                        {
                            mail.Attachments.Add(new Attachment(filename));
                            mail.Attachments.Add(new Attachment(namafile));
                        }
                    }

                    using (SmtpClient smtp = new SmtpClient())
                    {
                        smtp.UseDefaultCredentials = false;
                        smtp.EnableSsl = true;
                        smtp.Credentials = new System.Net.NetworkCredential("ujangng09@gmail.com", "fpowdjjbzsluqsae");
                        smtp.Host = "smtp.gmail.com";
                        smtp.Port = 587;
                        smtp.Send(mail);
                    };
                }
            }
            return View("Index", data);
        }
    }
}
