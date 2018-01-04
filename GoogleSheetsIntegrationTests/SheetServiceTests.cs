using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace Beardiegames.GoogleSheetsIntegration.Tests
{
    [TestClass()]
    public class SheetServiceTests
    {
        [TestMethod()]
        public void RunTest()
        {
            // Test1: Setup a new service
            SheetService service = new SheetService();
            // test if values are untouched
            Assert.IsNull(service.peekCredential);
            Assert.IsNull(service.peekRunningService);

            // Test2: Run the service
            service.Run("client_id.json", "SheetServiceTest RunTest");
            // test if the service and credential are set
            Assert.IsNotNull(service.peekCredential);
            Assert.IsNotNull(service.peekRunningService);
            // test expected values to be set
            Assert.AreEqual<string>(service.peekCredential.UserId, "user");
            Assert.AreSame(service.peekRunningService.HttpClientInitializer, service.peekCredential);
            Assert.AreEqual<string>(service.peekRunningService.ApplicationName, "SheetServiceTest RunTest");
        }

        [TestMethod()]
        public void ResourceTest()
        {
            // Prep: Setup a new service
            SheetService service = new SheetService();

            // Test1: Run the service
            service.Run("client_id.json", "SheetServiceTest ResourceTest");
            // test if the service and credential are set
            Assert.IsNotNull(service.peekCredential);
            Assert.IsNotNull(service.peekRunningService);

            // Test2: Get Resources
            SpreadsheetsResource resource = service.Resource();
            Assert.IsNotNull(resource);

            // Test3: Test if we can use the resource create a new 'get spreadsheet' request
            SpreadsheetsResource.GetRequest request = resource.Get("1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");
            Assert.IsNotNull(resource);
            Assert.AreEqual<string>(request.SpreadsheetId, "1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");

            // Test4: Execute the request and test the response data
            Spreadsheet response = request.Execute();
            Assert.IsNotNull(response);
            Assert.AreEqual<string>(response.SpreadsheetId, "1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");
        }
    }
}