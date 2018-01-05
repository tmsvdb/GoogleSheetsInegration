using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;


namespace Beardiegames.GoogleSheetsIntegration
{
    public class APIService
    {
        // Properties
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        UserCredential credential;
        SheetsService service;

        // Class information
        public UserCredential peekCredential { get { return credential; } }
        public SheetsService peekRunningService { get { return service; } }

        // public methodes

        public void Run(string cliend_id_location, string application_name)
        {
            credential = GetCredential(cliend_id_location);
            service = CreateAPIService(credential, application_name);
        }

        public SpreadsheetsResource Resource()
        {
            return service.Spreadsheets;
        }

        // Local Implementation

        private UserCredential GetCredential(string cliend_id_location)
        {
            UserCredential credential;

            using (var stream = new FileStream(cliend_id_location, FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            return credential;
        }

        private SheetsService CreateAPIService(UserCredential credential, string ApplicationName)
        {
            // Create Google Sheets API service.
            return new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }
    }
}