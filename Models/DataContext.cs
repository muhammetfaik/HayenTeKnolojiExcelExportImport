using Microsoft.EntityFrameworkCore;

namespace HayenTeKnolojiExcelExportImport.Models
{
    public class DataContext : DbContext
    {
        public DataContext(DbContextOptions<DataContext> opt) : base(opt)
        {
        }

        public DbSet<Product> Products { get; set; }
    }
}
