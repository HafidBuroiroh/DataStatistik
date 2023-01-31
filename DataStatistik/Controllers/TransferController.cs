using Microsoft.AspNetCore.Mvc;
using DataStatistik.Data;
using SqlConnection = Microsoft.Data.SqlClient.SqlConnection;
using SqlCommand = Microsoft.Data.SqlClient.SqlCommand;
using SqlDataReader = Microsoft.Data.SqlClient.SqlDataReader;
using SqlBulkCopy = Microsoft.Data.SqlClient.SqlBulkCopy;

namespace DataStatistik.Controllers
{
    public class TransferController : Controller
    {
        private readonly DataContext _context;
        public IConfiguration _Configuration { get; set; }
        public TransferController(IConfiguration configuration, DataContext context)
        {
            _Configuration = configuration;
            _context = context;
        }
        public IActionResult Index()
        {
            return View(_context.data_statistik.ToList());
        }
        public IActionResult Transfer(DateTime from, DateTime until, string btnsubmit)
        {
            var data = from d in _context.data_statistik select d;
            ViewData["CurrentFIlter"] = from;
            ViewData["CurrentFilters"] = until;

            if (btnsubmit == "Search")
            {
                if (from != null && until != null)
                {
                    data = data.Where(d => d.Date >= from && d.Date <= until);
                }
            }
            if (btnsubmit == "Transfer")
            {
                var connectionstring = _Configuration["ConnectionStrings:DefaultConnection"];
                SqlConnection con = new SqlConnection(connectionstring);
                var sql = $"SELECT * FROM data_statistik WHERE Date Between '{from.ToString("yyyy-MM-dd")}' AND '{until.ToString("yyyy-MM-dd")}'";
                SqlCommand com = new SqlCommand(sql, con);
                con.Open();

                SqlDataReader dataReader = com.ExecuteReader();

                SqlBulkCopy sqlBulk = new SqlBulkCopy(con);
                sqlBulk.DestinationTableName = "dbo.data_statistik_saved";

                sqlBulk.ColumnMappings.Add("Date", "Date");
                sqlBulk.ColumnMappings.Add("Period", "Period");
                sqlBulk.ColumnMappings.Add("MemberCode", "MemberCode");
                sqlBulk.ColumnMappings.Add("MemberName", "MemberName");
                sqlBulk.ColumnMappings.Add("Province", "Province");
                sqlBulk.ColumnMappings.Add("City", "City");
                sqlBulk.ColumnMappings.Add("Frequency", "Frequency");
                sqlBulk.ColumnMappings.Add("Volume", "Volume");
                sqlBulk.ColumnMappings.Add("Value", "Value");


                sqlBulk.WriteToServer(dataReader);
            }
            return View("Index", data);
        }
    }
}
