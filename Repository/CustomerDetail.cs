using Microsoft.Data.SqlClient;
using System.Data;
using System.Data.OleDb;

namespace HayenTeKnolojiExcelExportImport.Repository
{
    public class CustomerDetail : ICustomer
    {
        private IConfiguration configuration;
        private IWebHostEnvironment webHostEnvironment;

        public CustomerDetail(IConfiguration configuration,IWebHostEnvironment webHostEnvironment)
        {
            this.configuration = configuration;
            this.webHostEnvironment = webHostEnvironment;
        }

        public DataTable CustomerDataTable(string path)
        {
            var conStr = configuration.GetConnectionString("excelconnection");
            DataTable dataTable = new DataTable();
            conStr = string.Format(conStr, path);
            using (OleDbConnection excelconn = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand()) { 
                using (OleDbDataAdapter adapterexcel = new OleDbDataAdapter())
                {
                    excelconn.Open();
                    cmd.Connection = excelconn;
                    DataTable excelschema;
                    excelschema = excelconn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    var sheetname = excelschema.Rows[0]["Table_Name"].ToString();
                    excelconn.Close();
                    excelconn.Open();
                    cmd.CommandText = "Select * from [" + sheetname + "]";
                    adapterexcel.SelectCommand = cmd;
                    adapterexcel.Fill(dataTable);
                    excelconn.Close();
                }
                }

            }
            return dataTable;
        }

       

        public string DocumentUpload(IFormFile formFile)
        {
            string uploadPath = webHostEnvironment.WebRootPath;
            string dest_path = Path.Combine(uploadPath, "uploaded_doc");
            if(!Directory.Exists(dest_path))
            {
                Directory.CreateDirectory(dest_path);
            }
            string sourceFile = Path.GetFileName(formFile.FileName);
            string path = Path.Combine (dest_path, sourceFile);
            using (FileStream fileStream = new FileStream(path,FileMode.Create))
            {
                formFile.CopyTo(fileStream);
            }
            return path;
        }

        public void ImportCustomer(DataTable customer)
        {
            var sqlconn = configuration.GetConnectionString("sqlconnection");
            using(SqlConnection scon = new SqlConnection(sqlconn))
            {
                using(SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(scon))
                {
                    sqlBulkCopy.DestinationTableName = "Customers";
                    sqlBulkCopy.ColumnMappings.Add("FirstName", "firstName");
                    sqlBulkCopy.ColumnMappings.Add("LastName", "lastName");
                    sqlBulkCopy.ColumnMappings.Add("Job", "job");
                    sqlBulkCopy.ColumnMappings.Add("Amount", "amount");
                    sqlBulkCopy.ColumnMappings.Add("TDate", "tdate");
                    scon.Open();
                    sqlBulkCopy.WriteToServer(customer);
                    scon.Close();
                }
            }
        }   
    }
}
