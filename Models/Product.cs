namespace HayenTeKnolojiExcelExportImport.Models
{
    public class Product
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
        public int Quantity { get; set; }
        public double Price { get; set; }
        public bool IsActive { get; set; }
        public DateTime ExpiryDate { get; set; }

    }
}
