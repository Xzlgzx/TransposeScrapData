using System.Data;
using HtmlAgilityPack;
using ExcelDataReader;

namespace ParseDownload
{
    internal class Program
    {
        static async Task Main(string[] args) // Main function that controls all actions
        {
            await DownloadPage.GetLinkHtml("https://www.abs.gov.au/statistics/" +
                                           "labour/employment-and-unemployment/labour-force-australia");
            DownloadPage.ScrapeLatestReleaseLink();
            await DownloadPage.GetExcelFile();
            DownloadPage.CreateCsv();
        }
    }

    internal static class DownloadPage
    {
        private static string _latestReleaseLink;
        private const string Baseurl = "https://www.abs.gov.au";
        private static string _xlsxBase = "/6202001.xlsx"; // This is the same across all dates, so I assumed it will not change.
        private static HtmlDocument _latestPgHtml;
        private static string _downloadFilePath;

        public static async Task GetLinkHtml(string url)
        {
            using HttpClient client = new HttpClient();
            
            // Pretend we are using a browser to bypass security
            client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/5" +
                                                           "37.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3");
            
            var response = "";
            
            try
            {
                response = await client.GetStringAsync(url);
            }
            catch (HttpRequestException e)
            {
                Console.WriteLine($"Request error: {e.Message}");
            }
            
            _latestPgHtml = new HtmlDocument();
            _latestPgHtml.LoadHtml(response);
        }
        
        public static void ScrapeLatestReleaseLink()
        {
            var linkNode = _latestPgHtml.DocumentNode.SelectSingleNode("//div[@id='content']//a");
                
            if (linkNode != null)
            {
                _latestReleaseLink = linkNode.GetAttributeValue("href", string.Empty);
                Console.WriteLine("The latestReleaseLink is: " + _latestReleaseLink);
            }
            else // Simplified for now
            {
                throw new KeyNotFoundException("We could not locate the URL. Possibly because HTML structure of pg has changed");
            }
        }
        
        public static async Task GetExcelFile()
        {
            // URL of the file to be downloaded
            string fileUrl = Baseurl + _latestReleaseLink + _xlsxBase;
        
            // Define the folder and file name
            string downloadFolderName = "Downloads";
            string fileName = "data.xlsx";
        
            // Get the current directory of the application
            string currentDirectory = Directory.GetCurrentDirectory(); // Potentially problematic due to system setting
        
            // Construct the full path to the downloads folder
            string downloadsFolderPath = Path.Combine(currentDirectory, downloadFolderName); 
        
            // Ensure the downloads folder exists
            if (!Directory.Exists(downloadsFolderPath))
            {
                Directory.CreateDirectory(downloadsFolderPath);
            }
        
            // Construct the full file path
            _downloadFilePath = Path.Combine(downloadsFolderPath, fileName);
        
            // Download the file
            using (HttpClient client = new HttpClient())
            {
                
                // Pretend we are using a browser to bypass security again
                client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/5" +
                                                               "37.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3");

                try
                {
                    Console.WriteLine("Starting download...");
                    byte[] fileBytes = await client.GetByteArrayAsync(fileUrl);
                    Console.WriteLine("Download completed.");
                
                    // Save the file to the specified path
                    await File.WriteAllBytesAsync(_downloadFilePath, fileBytes);
                    Console.WriteLine($"File saved to {_downloadFilePath}");
                }
                catch (HttpRequestException e)
                {
                    Console.WriteLine($"Request error: {e.Message}");
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Unexpected error: {e.Message}");
                }
            }
        }
        
        public static void CreateCsv()
        {
            string outputPath = "transposed.csv"; // Output CSV file

            // Register the encoding provider
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            
            using (var stream = File.Open(_downloadFilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    var dataTable = result.Tables["Data1"];

                    // Find the starting row
                    int startRow = -1;

                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        if (dataTable.Rows[i][0]?.ToString() == "Series ID")
                        {
                            startRow = i;
                            break;
                        }
                    }

                    // If "Series ID" was found
                    if (startRow != -1)
                    {
                        using (var writer = new StreamWriter(outputPath))
                        {
                            // Create a new DataTable for the transposed data
                            DataTable transposedData = new DataTable();

                            // Add columns to the transposed DataTable
                            for (int row = startRow; row < dataTable.Rows.Count; row++)
                            {
                                transposedData.Columns.Add(dataTable.Rows[row][0].ToString());
                            }

                            // Read rows and transpose
                            for (int col = 0; col < dataTable.Columns.Count; col++)
                            {
                                // Create a new row for each column in the original data
                                var newRow = transposedData.NewRow();
                                for (int row = startRow; row < dataTable.Rows.Count; row++)
                                {
                                    if (row > startRow && col == 0) // Case 1: formatting columns
                                    {
                                        string formattedDate = ((DateTime) dataTable.Rows[row][col]).ToString("MMM-yy", System.Globalization.CultureInfo.InvariantCulture);
                                        newRow[row - startRow] = formattedDate;
                                    }
                                    else if (row > startRow) // Case 2: round numbers
                                    {
                                        newRow[row - startRow] = Math.Round((Double) dataTable.Rows[row][col], 1);
                                    }
                                    else // Case 3: 1st column
                                    {
                                        newRow[row - startRow] = dataTable.Rows[row][col];
                                    }
                                }
                                writer.WriteLine(string.Join(",", newRow.ItemArray)); // Write line
                            }
                        }
                    }
                    else
                    {
                        throw new KeyNotFoundException("Cannot find Series ID in column 1");
                    }
                }
            }
        }
    }
}