using System.IO;
using System.Reflection.PortableExecutable;
using Syncfusion.HtmlConverter;
using Syncfusion.Pdf;
using MsgReader;

namespace Test_HTML_to_PDF
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        private FileSystemWatcher _fileSystemWatcher;
        private string sourceFolderPath = @"E:\HD Bundle\Original Files"; // Set the folder you want to watch
        private string destinationFolderPath = @"E:\HD Bundle\Converted Files"; // Set the folder you want to copy files to

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }

        public override Task StartAsync(CancellationToken cancellationToken)
        {
            _fileSystemWatcher = new FileSystemWatcher(sourceFolderPath);
            _fileSystemWatcher.Filter = "*.msg, *.eml"; // Watch for email file types
            _fileSystemWatcher.Created += OnNewFileDetected;
            _fileSystemWatcher.EnableRaisingEvents = true;

            _logger.LogInformation("FolderWatcherService started.");
            return base.StartAsync(cancellationToken);
        }

        protected override Task ExecuteAsync(CancellationToken stoppingToken)
        {
            // The service runs indefinitely
            return Task.CompletedTask;
        }

        private void OnNewFileDetected(object sender, FileSystemEventArgs e)
        {
            Task.Run(() => CopyAndDeleteFileAsync(e.FullPath));
        }

        private async Task CopyAndDeleteFileAsync(string filePath)
        {
            try
            {
                string fileName = Path.GetFileName(filePath);
                string fileType = Path.GetExtension(filePath).ToLower();
                string origFileType = Path.GetExtension(filePath);
                fileName = fileName.Replace(origFileType, ".pdf");
                string destFilePath = Path.Combine(destinationFolderPath, fileName);

                //_logger.LogCritical("Found file");
                await WaitForFileAsync(filePath);

                if (fileType == ".msg" || fileType == ".eml")
                {
                    await ConvertEmail(filePath, destFilePath);
                }
                await Task.Delay(1000);
                await WaitForFileAsync(filePath);
                File.Delete(filePath);
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        private async Task WaitForFileAsync(string filePath)
        {
            bool fileIsReady = false;
            while (!fileIsReady)
            {
                try
                {
                    using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        fileIsReady = true;
                    }

                }
                catch (IOException)
                {
                    await Task.Delay(100);
                }
            }
        }

        private async Task ConvertEmail(string filePath, string destinationPath)
        {

            HtmlToPdfConverter htmlConverter = new HtmlToPdfConverter();
            string fullPath = System.IO.Path.GetFullPath(filePath);
            //Convert EML or MSG to HTML using a third-party reader.
            var msgReader = new Reader();

            string tempFolder = @"E:\HD Bundle\Temp";
            //File contains the HTML file converted from MSG.
            var files = msgReader.ExtractToFolder(fullPath, tempFolder);

            var error = msgReader.GetErrorMessage();

            if (!string.IsNullOrEmpty(error))

                throw new Exception(error);

            if (!string.IsNullOrEmpty(files[0]))
            {
                htmlConverter = new HtmlToPdfConverter();
                BlinkConverterSettings settings = new BlinkConverterSettings();
                //Assign Blink converter settings to HTML converter.
                htmlConverter.ConverterSettings = settings;
            }
            //Convert HTML file to a PDF document. 
            PdfDocument pdfDocument = htmlConverter.Convert(System.IO.File.ReadAllText(files[0]), tempFolder);
            using (FileStream outputFileStream = new FileStream(destinationPath, FileMode.Create, FileAccess.ReadWrite))
            {
                //Save the PDF document to file stream.
                pdfDocument.Save(outputFileStream);
            }
            //Close the document.
            pdfDocument.Close(true);
        }
    }
}
