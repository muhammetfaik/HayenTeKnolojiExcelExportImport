using Microsoft.EntityFrameworkCore;
namespace HayenTeKnolojiExcelExportImport.Models
{
    public class CustDBContext:DbContext
    {
        public CustDBContext(DbContextOptions<CustDBContext> options) : base(options) {
                
        }

        public virtual DbSet<Customer> Customers { get; set; }


    }
}
