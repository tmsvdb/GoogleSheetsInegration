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
        public void DocumentTest()
        {
            APIService service = new APIService();
            service.Run("client_id.json", "SheetTests ConstructorTest");
            
            Document doc = new Document(service.Resource(), "1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");
            Assert.AreEqual<string>(doc.Id, "1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");
            Assert.IsNotNull(doc.lookupResource);
        }

        [TestMethod()]
        public void ReadValuesTest()
        {
            APIService service = new APIService();
            service.Run("client_id.json", "SheetTests ReadValuesTest");
            Document doc = new Document(service.Resource(), "1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");

            doc.ReadSpreadsheet();

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

            //OutputValueList("# Read Values", values);*/
        }

        [TestMethod()]
        public void WriteValuesTest()
        {
            APIService service = new APIService();
            service.Run("client_id.json", "SheetTests ReadValuesTest");
            Document doc = new Document(service.Resource(), "1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");
            
            // get origional values from sheet
            doc.ReadSpreadsheet();

            IList<IList<object>> ori_values = Page.ToObjectList(doc.GetPageByTitle("TestPage1"));

            //OutputValueList("# Origional Values1", ori_values);
            //OutputValueList("# Origional Values2", Page.ToObjectList(doc.GetPageByTitle("TestPage2")));

            Page p = doc.GetPageByTitle("TestPage1");
            p.SetValue(4, 0, "Test col A");
            p.SetValue(4, 1, "Test col B");
            p.SetValue(4, 2, "Test col C");

            IList<IList<object>> write_values = Page.ToObjectList(doc.GetPageByTitle("TestPage1"));

            //OutputValueList("# Write Values", write_values);

            doc.WriteSpreadsheet();

            IList<IList<object>> response_values = Page.ToObjectList(doc.GetPageByTitle("TestPage1"));
            Assert.IsNotNull(response_values);
            //OutputValueList("# Response Values1", response_values);
            //OutputValueList("# Response Values2", Page.ToObjectList(doc.GetPageByTitle("TestPage2")));

            // loop[ through response values and compare them with the write_values
            for (int row = 0; row < response_values.Count; row++)
                for (int col = 0; col < response_values[row].Count; col++)
                    Assert.AreEqual(write_values[row][col], response_values[row][col]);


            doc.SetPageByTitle("TestPage1", Page.FromObjectList("TestPage1", ori_values));
            doc.WriteSpreadsheet();

            // Test2: reset values to the origional values
            //OutputValueList("# Reset Values1", Page.ToObjectList(doc.GetPageByTitle("TestPage1")));
            //OutputValueList("# Reset Values2", Page.ToObjectList(doc.GetPageByTitle("TestPage2")));
        }

        private void OutputValueList (string listname, IList<IList<object>> values)
        {
            Console.WriteLine(listname);

            for (int row = 0; row < values.Count; row++)
            {
                string line = "";
                for (int col = 0; col < values[row].Count; col++)
                    line += values[row][col] + (col < values[row].Count - 1 ? "," : "");
                Console.WriteLine(line);
            }
        }
    }
}