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
    public class DocumentTests
    {
        [TestMethod()]
        public void Document_DocumentTest()
        {
            APIService service = new APIService();
            service.Run("client_id.json", "SheetTests ConstructorTest");

            Document doc = new Document(service.Resource(), "1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");
            Assert.AreEqual<string>(doc.Id, "1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");
            Assert.IsNotNull(doc.lookupResource);
        }

        [TestMethod()]
        public void Document_GetPageByTitleTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void Document_SetPageByTitleTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void Document_ReadSpreadsheetTest()
        {
            APIService service = new APIService();
            service.Run("client_id.json", "SheetTests ReadValuesTest");
            Document doc = new Document(service.Resource(), "1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");

            if (!doc.ReadSpreadsheet())
                Assert.Fail();

            Assert.AreEqual<int>(doc.numberOfPages, 2);

            Page p1 = doc.GetPageByTitle("TestPage1");
            Page p2 = doc.GetPageByTitle("TestPage2");
            Page p3 = doc.GetPageByTitle("TestPage3");

            Assert.IsNotNull(p1);
            Assert.IsNotNull(p2);
            Assert.IsNull(p3);

            Assert.AreEqual<int>(p1.peekHeight, 5);
            Assert.AreEqual<int>(p1.peekWidth, 3);
            Assert.AreEqual<int>(p2.peekHeight, 1);
            Assert.AreEqual<int>(p2.peekWidth, 2);

            Assert.AreEqual(p1.GetValue(0, 0), "project");
            Assert.AreEqual(p1.GetValue(0, 1), "id");
            Assert.AreEqual(p1.GetValue(0, 2), "status");

            Assert.AreEqual(p2.GetValue(0, 0), "Hello");
            Assert.AreEqual(p2.GetValue(0, 1), "World");
        }

        [TestMethod()]
        public void Document_WriteSpreadsheetTest()
        {
            APIService service = new APIService();
            service.Run("client_id.json", "SheetTests ReadValuesTest");

            Document origional_doc = new Document(service.Resource(), "1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");
            Document test_doc = new Document(service.Resource(), "1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");

            // get origional values from sheet
            if (!origional_doc.ReadSpreadsheet()) Assert.Fail();
            if (!test_doc.ReadSpreadsheet()) Assert.Fail();

            Page p1 = test_doc.GetPageByTitle("TestPage1");
            p1.SetValue("A5", "Test col A");
            p1.SetValue("B5", "Test col B");
            p1.SetValue("C5", "Test col C");

            if (!test_doc.WriteSpreadsheet())
                Assert.Fail();

            Page p2 = test_doc.GetPageByTitle("TestPage1");
            Assert.AreEqual<string>("Test col A", p2.GetValue("A5"));
            Assert.AreEqual<string>("Test col B", p2.GetValue("B5"));
            Assert.AreEqual<string>("Test col C", p2.GetValue("C5"));

            if (!origional_doc.WriteSpreadsheet())
                Assert.Fail();

            Page p3 = test_doc.GetPageByTitle("TestPage1");
            Assert.AreEqual<string>("lastline", p3.GetValue("A5"));
            Assert.AreEqual<string>("lastid", p3.GetValue("B5"));
            Assert.AreEqual<string>("laststatus", p3.GetValue("C5"));
        }

        [TestMethod()]
        public void Document_CreateNewDocumentTest()
        {
            APIService service = new APIService();
            service.Run("client_id.json", "Document_CreateNewDocumentTest");

            Document doc = Document.CreateNewDocument(service.Resource(), "new test document");

            Assert.IsNotNull(doc);
            Assert.IsTrue(doc.ReadSpreadsheet());
            Assert.AreEqual<string>("new test document", doc.peekProperties.Title);
        }
    }
}