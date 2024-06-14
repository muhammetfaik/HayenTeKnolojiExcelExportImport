using Microsoft.EntityFrameworkCore.Storage.ValueConversion.Internal;

namespace HayenTeKnolojiExcelExportImport.Models
{
    public class ProductRequest
    {
        public string Name { get; set; }
        public int Quantity { get; set; }
        public double Price { get; set; }
        public bool IsActive { get; set; }
        public DateTime ExpiryDate { get; set; }
    }
}
