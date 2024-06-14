using HayenTeKnolojiExcelExportImport.Models;
using HayenTeKnolojiExcelExportImport.Repository;
using Microsoft.EntityFrameworkCore;

namespace HayenTeKnolojiExcelExportImport
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);
            builder.Logging.ClearProviders();
            builder.Logging.AddConsole();
            // Add services to the container.
            builder.Services.AddControllersWithViews();
            builder.Services.AddDbContext<CustDBContext>(conn => conn.UseSqlServer(builder.Configuration.GetConnectionString("sqlconnection")));
            builder.Services.AddScoped<ICustomer, CustomerDetail>();
            builder.Services.AddDbContext<DataContext>(conn => conn.UseSqlServer(builder.Configuration.GetConnectionString("sqlconnection")));
            var app = builder.Build();

            // Configure the HTTP request pipeline.
            if (!app.Environment.IsDevelopment())
            {
                app.UseExceptionHandler("/Home/Error");
                // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
                app.UseHsts();
            }

            app.UseHttpsRedirection();
            app.UseStaticFiles();

            app.UseRouting();

            app.UseAuthorization();

            app.MapControllerRoute(
                name: "default",
                pattern: "{controller=Home}/{action=Index}/{id?}");

            app.Run();
        }
    }
}
