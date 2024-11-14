using Test_HTML_to_PDF;

namespace Test_HTML_to_PDF
{
    public class Program
    {
        public static void Main(string[] args)
        {
            CreateHostBuilder(args).Build().Run();
        }

        public static IHostBuilder CreateHostBuilder(string[] args) =>
            Host.CreateDefaultBuilder(args)
                .ConfigureServices((hostContext, services) =>
                {
                    services.AddHostedService<Worker>();
                    Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("");
                })
                .UseWindowsService(); // Makes it run as a Windows Service

    }
}
