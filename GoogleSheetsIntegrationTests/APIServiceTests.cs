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

using Beardiegames.GoogleSheetsIntegration;

namespace Beardiegames.GoogleSheetsIntegration.Tests
{
    [TestClass()]
    public class APIServiceTests
    {
        [TestMethod()]
        public void APIService_RunTest()
        {
            APIService service = new APIService();
            Assert.IsNull(service.peekCredential);
            Assert.IsNull(service.peekRunningService);

            service.Run("client_id.json", "SheetServiceTest RunTest");
            Assert.IsNotNull(service.peekCredential);
            Assert.IsNotNull(service.peekRunningService);
            Assert.AreEqual<string>(service.peekCredential.UserId, "user");
            Assert.AreSame(service.peekRunningService.HttpClientInitializer, service.peekCredential);
            Assert.AreEqual<string>(service.peekRunningService.ApplicationName, "SheetServiceTest RunTest");
        }

        [TestMethod()]
        public void APIService_ResourceTest()
        {
            APIService service = new APIService();

            service.Run("client_id.json", "SheetServiceTest ResourceTest");
            Assert.IsNotNull(service.peekCredential);
            Assert.IsNotNull(service.peekRunningService);

            SpreadsheetsResource resource = service.Resource();
            Assert.IsNotNull(resource);

            SpreadsheetsResource.GetRequest request = resource.Get("1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");
            Assert.IsNotNull(resource);
            Assert.AreEqual<string>(request.SpreadsheetId, "1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");

            Spreadsheet response = request.Execute();
            Assert.IsNotNull(response);
            Assert.AreEqual<string>(response.SpreadsheetId, "1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");
        }
    }
}