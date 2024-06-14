using HayenTeKnolojiExcelExportImport.Models;
using HayenTeKnolojiExcelExportImport.Repository;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace HayenTeKnolojiExcelExportImport.Controllers
{
    public class ProductController : Controller
    {
        private readonly DataContext _context;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private readonly ILogger<ProductController> _logger;


        public ProductController(DataContext context, IWebHostEnvironment webHostEnvironment,ILogger<ProductController> logger)
        {
            _context = context;
            _webHostEnvironment = webHostEnvironment;
            _logger = logger;
            _logger.Log(LogLevel.Information, "Hello from Product Controller");
        }

        [HttpPost("Upload")]
        [DisableRequestSizeLimit]
        public async Task<ActionResult> Upload(CancellationToken ct)
        {
            if (Request.Form.Files.Count == 0) return NoContent();

            var file = Request.Form.Files[0];
            var filePath = SaveFile(file);

            // load product requests from excel file
            var productRequests = ExcelHelper.Import<ProductRequest>(filePath);

            // save product requests to database
            foreach (var productRequest in productRequests)
            {
                var product = new Product
                {
                    Id = Guid.NewGuid(),
                    Name = productRequest.Name,
                    Quantity = productRequest.Quantity,
                    Price = productRequest.Price,
                    IsActive = productRequest.IsActive,
                    ExpiryDate = productRequest.ExpiryDate
                };
                await _context.AddAsync(product, ct);
            }
            await _context.SaveChangesAsync(ct);

            return Ok();
        }

        private string SaveFile(IFormFile file) {

            if (file.Length == 0)
            {
                throw new BadHttpRequestException("File is empty.");
            }

            var extension = Path.GetExtension(file.FileName);

            var webRootPath = _webHostEnvironment.WebRootPath;
            if (string.IsNullOrWhiteSpace(webRootPath))
            {
                webRootPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
            }

            var folderPath = Path.Combine(webRootPath, "uploads");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            var fileName = $"{Guid.NewGuid()}{extension}";
            var filePath = Path.Combine(folderPath, fileName);
            using var stream = new FileStream(filePath, FileMode.Create);
            file.CopyTo(stream);

            return filePath;
        }

        [HttpGet("Download")]
        public async Task<FileResult> Download(CancellationToken ct)
        {
            var products = await _context.Products.ToListAsync(ct);
            var file = ExcelHelper.CreateFile(products);
            return File(file, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "products.xlsx");
        }

        [HttpGet("createdata")]
        public async void CreateData()
        {
            _context.Products.Add(new Product
            {
                Id = Guid.NewGuid(),
                Name = "Product 1",
                Quantity = 10,
                Price = 100,
                IsActive = true,
                ExpiryDate = DateTime.Now.AddDays(10)
            });

            _context.Products.Add(new Product
            {
                Id = Guid.NewGuid(),
                Name = "Product 2",
                Quantity = 20,
                Price = 200,
                IsActive = true,
                ExpiryDate = DateTime.Now.AddDays(20)
            });

            _context.SaveChanges();
        }

        public IActionResult Index()
        {
            return View();
        }
    }
}
