using ClosedXML.Excel;
using HayenTeKnolojiExcelExportImport.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Microsoft.Data.SqlClient;
using System.Data;
using HayenTeKnolojiExcelExportImport.Repository;

namespace HayenTeKnolojiExcelExportImport.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private IWebHostEnvironment _webHostEnvironment;
        private readonly ICustomer _customer;
        List<Customer> customers = new List<Customer>();
        SqlCommand com = new SqlCommand();
        SqlDataReader dr;
        SqlConnection con = new SqlConnection();

        public HomeController(ILogger<HomeController> logger,IWebHostEnvironment  webHostEnvironment,ICustomer customer)
        {
            _logger = logger;
            _webHostEnvironment = webHostEnvironment;
            _customer = customer;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Index(IFormFile formFile)
        {
            _logger.LogInformation("This is upload file");
            string path = _customer.DocumentUpload(formFile);
            DataTable dt = _customer.CustomerDataTable(path);
            _customer.ImportCustomer(dt);
            return View();
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


        public IActionResult ExportExcel()
        {
            using (var workbook = new XLWorkbook())
            {
                _logger.LogInformation("Export Excel Started");
                var ws = workbook.Worksheets.Add("Student");
                ws.Range("A2:E2").Merge();
                ws.Cell(1,1).Value = "Report";
                ws.Cell(1,1).Style.Font.Bold = true;
                ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(1, 1).Style.Font.FontSize = 30;


                // Header 
                ws.Cell(4, 1).Value = "DocEntry";
                ws.Cell(4, 2).Value = "Name";
                ws.Cell(4, 1).Value = "Gender";
                ws.Cell(4, 1).Value = "Phone";
                ws.Cell(4, 1).Value = "Email";
                ws.Range("A4:E4").Style.Fill.BackgroundColor = XLColor.Alizarin;

                //Connect to SQL here
                DataTable dt = new DataTable();
                //SqlConnection con = new SqlConnection("Server=(localdb)\\MSSQLLocalDB;Database=SportsStore;MultipleActiveResultSets=true");
                SqlConnection con_two = new SqlConnection("Server=MUHAMMETFAIK;Database=Northwind;Integrated Security=True;Trust Server Certificate=True;");
                SqlDataAdapter ad_two = new SqlDataAdapter("select * from dbo.Shippers",con_two);
                //SqlDataAdapter ad = new SqlDataAdapter("select * from dbo.Products", con);
                ad_two.Fill(dt);
                int i = 5;
                foreach(DataRow row in dt.Rows)
                {
                    ws.Cell(i, 1).Value = row[0].ToString();
                    ws.Cell(i, 2).Value = row[1].ToString();
                    ws.Cell(i, 3).Value = row[2].ToString();
                    //ws.Cell(i, 4).Value = row[3].ToString();
                    i = i + 1;
                }
                i = i - 1;
                ws.Cells("A4:E" + 1).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                ws.Cells("A4:E" + 1).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                ws.Cells("A4:E" + 1).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                ws.Cells("A4:E" + 1).Style.Border.RightBorder = XLBorderStyleValues.Thin;

                using(var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(content,"application/vnd.openxmlformats-officedocument-spreadsheetml.sheet","Student.xlsx");
                }
            };
        }

        /**public IActionResult ImportExcel()
        {

        }**/

        public void Sqlconnectioncontrol() { 
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection("Server=(localdb)\\MSSQLLocalDB;Database=SportsStore;MultipleActiveResultSets=true");
            SqlDataAdapter ad = new SqlDataAdapter("select * from dbo.Products", con);
           
            ad.Fill(dt);

            for(int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow row = dt.Rows[i];
            }
            
        }

        //Gösterim için bir kod.Sonra anlatýlacak.
        /*
        private void FetchData()
        {
            if(customers.Count > 0)
            {
                customers.Clear();
            }
            try
            {
                con.Open();
                com.Connection = con;
                com.CommandText = "SELECT * FROM dbo.Customers";
                dr = com.ExecuteReader();
                while(dr.Read())
                {
                    customers.Add(new Customer() { id = Convert.ToInt32(dr["id"]),
                        firstName = dr["firstName"].ToString(),
                        lastName = dr["lastName"].ToString(),
                        job = dr["job"].ToString(),
                        amount = Convert.ToSingle(dr["amount"]),
                        tdate = Convert.ToDateTime(dr["tdate"] });


                }

            }
            catch(Exception ex)
            {
                throw ex;
            }
        }*/
    }
}
