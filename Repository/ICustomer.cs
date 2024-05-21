using System.Data;

namespace HayenTeKnolojiExcelExportImport.Repository
{
    public interface ICustomer
    {
        string DocumentUpload(IFormFile formFile);
        DataTable CustomerDataTable(string path);
        void ImportCustomer(DataTable customer);

    }
}
