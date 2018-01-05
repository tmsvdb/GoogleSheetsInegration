using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Beardiegames.GoogleSheetsIntegration;

namespace Beardiegames.GoogleSheetsIntegration.Tests
{
    [TestClass()]
    public class SheetManagerTests
    {
        [TestMethod()]
        public void SheetManagerTest()
        {
            DocumentManager manager = new DocumentManager("client_id.json", "Test SheetManager");
            Assert.IsTrue(manager.isRunning);
        }

        [TestMethod()]
        public void GetSheetTest()
        {
            DocumentManager manager = new DocumentManager("client_id.json", "Test GetSheet");

            Document document1 = manager.GetSheet("1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");
            Assert.IsNotNull(document1);
            Assert.AreEqual<int>(manager.numberOfDocuments, 1);
            Assert.AreEqual<string>(document1.Id, "1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");

            Document document2 = manager.GetSheet("1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");
            Assert.IsNotNull(document2);
            Assert.AreEqual<int>(manager.numberOfDocuments, 1);
            Assert.AreSame(document1, document2);

            Document document3 = manager.GetSheet("1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms");
            Assert.IsNotNull(document3);
            Assert.AreEqual<int>(manager.numberOfDocuments, 2);
            Assert.AreNotSame(document1, document3);
            Assert.AreEqual<string>(document3.Id, "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms");
        }
    }
}