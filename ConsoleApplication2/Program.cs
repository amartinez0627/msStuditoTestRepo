using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Download;

namespace ConsoleApplication2
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        

        static void Main(string[] args)
        {
            //getInfoFromExcel();
            getInfoFromDrive();
           // testCode();
            Console.Read();
        }

        public static void getInfoFromExcel (){

             string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
             string ApplicationName = "Google Sheets API .NET Quickstart";

            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.
            String spreadsheetId = "12wChw6ce9gUqcXg_akJ8zmlyAwbDO4H9aTY2I56AQC8";
            String range = "Form Responses 1!A2:F";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            // Prints the names and majors of students in a sample spreadsheet:
            // https://docs.google.com/spreadsheets/d/12wChw6ce9gUqcXg_akJ8zmlyAwbDO4H9aTY2I56AQC8/edit
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {

                foreach (var row in values)
                {
                    // Print columns A and E, which correspond to indices 0 and 4.
                    for (int x = 0; x < row.Count; x++)
                    {
                        Console.Write("{0}    ", row[x]);
                    }

                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
        }

        public static void getInfoFromDrive()
        {

            string[] Scopes = { DriveService.Scope.Drive };
            string ApplicationName = "Google Drive API";

           // UserCredential credential;
            GoogleCredential credential;

            using (var stream =
                new FileStream("client_id.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";


            //    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
            //        GoogleClientSecrets.Load(stream).Secrets,
            //        Scopes,
            //        "user",
            //        CancellationToken.None,
            //        new FileDataStore(credPath, true)).Result;
            //    Console.WriteLine("Credential file saved to: " + credPath);
            }
                credential = GoogleCredential.FromFile(".\\client_id.json").CreateScoped(DriveService.Scope.Drive);

            // Create Drive API service.
            var driveService = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            //18daqA8U9E7rktS_yo8jTZZ1KqU9UdOMY
            var fileId = "1OrPAfEYMmVloR9_Djg7Da0cLHdsBoN3G";
            var request = driveService.Files.Get(fileId);
            var strea = new System.IO.MemoryStream();

            // Add a handler which will be notified on progress changes.
            // It will notify on each chunk download and when the
            // download is completed or failed.

            try
            {
                request.MediaDownloader.ProgressChanged +=
                (IDownloadProgress progress) =>
                {
                    switch (progress.Status)
                    {
                        case DownloadStatus.Downloading:
                            {
                                Console.WriteLine(progress.BytesDownloaded);
                                break;
                            }
                        case DownloadStatus.Completed:
                            {
                                Console.WriteLine("Download complete.");
                                DirectoryInfo di = new DirectoryInfo(".\\cv");
                                if (!di.Exists)
                                {
                                    di.Create();
                                }
                                using (System.IO.FileStream file = new System.IO.FileStream(".\\cv\\data.docx", System.IO.FileMode.Create, System.IO.FileAccess.Write))
                                {
                                    strea.WriteTo(file);
                                }
                                break;
                            }
                        case DownloadStatus.Failed:
                            {
                                Console.WriteLine("Download failed.");
                                break;
                            }
                    }
                };
                request.Download(strea);


                
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            

        }

        public static void testCode()
        {

         string[] Scopes = { DriveService.Scope.Drive };
         string ApplicationName = "Drive API";
            GoogleCredential credential;

            using (var stream =
                new FileStream("client_id.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                //string credPath = "token2.json";

                string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/drive-dotnet-quickstart.json");

                //credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                //    GoogleClientSecrets.Load(stream).Secrets,
                //    Scopes,
                //    "user",
                //    CancellationToken.None,
                //    new FileDataStore(credPath, true)).Result;

                credential = GoogleCredential.FromFile("c:\\users\\antonio_martinez\\documents\\visual studio 2013\\Projects\\ConsoleApplication2\\ConsoleApplication2\\client_id.json").CreateScoped(DriveService.Scope.Drive);

               // Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Drive API service.
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define parameters of request.
            FilesResource.ListRequest listRequest = service.Files.List();
            listRequest.PageSize = 1;
            listRequest.Fields = "nextPageToken, files(id, name)";

            // List files.
            IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute()
                .Files;
            Console.WriteLine("Files:");
            if (files != null && files.Count > 0)
            {
                foreach (var file in files)
                {
                    Console.WriteLine("{0} ({1})", file.Name, file.Id);
                }
            }
            else
            {
                Console.WriteLine("No files found.");
            }
            Console.Read();
        }
    }
}
