using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beardiegames.GoogleSheetsIntegration.Tests
{
    [TestClass()]
    public class SheetTests
    {
        [TestMethod()]
        public void SheetTest()
        {
            // Prep: pepare a running service to pass to the
            SheetService service = new SheetService();
            service.Run("client_id.json", "SheetTests ConstructorTest");
            
            //Test1: create a new sheet and test passed parameters
            Sheet sheet = new Sheet(service.Resource(), "1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");
            Assert.AreEqual<string>(sheet.Id, "1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");
            Assert.IsNotNull(sheet.lookupResource);
        }

        [TestMethod()]
        public void ReadValuesTest()
        {
            // Prep: pepare a running service to pass to the
            SheetService service = new SheetService();
            service.Run("client_id.json", "SheetTests ReadValuesTest");
            // create a new sheet
            Sheet sheet = new Sheet(service.Resource(), "1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");

            // Test1: 
            IList<IList<object>> values = sheet.ReadValues("Sheet1!A1:C5");
            Assert.IsNotNull(values);
            Assert.AreEqual<int>(values.Count, 5);
            Assert.AreEqual<int>(values[0].Count, 3);

            Assert.AreEqual(values[0][0], "project");
            Assert.AreEqual(values[0][1], "id");
            Assert.AreEqual(values[0][2], "status");

            OutputValueList("# Read Values", values);
        }

        [TestMethod()]
        public void WriteValuesTest()
        {
            // Prep: pepare a running service to pass to the
            SheetService service = new SheetService();
            service.Run("client_id.json", "SheetTests ReadValuesTest");
            // create a new sheet
            Sheet sheet = new Sheet(service.Resource(), "1E_Um33Mj8oIylVgqUIbiAfw6YH1YlpZy73bCIMWl-dc");
            // get origional values from sheet
            IList<IList<object>> ori_values = sheet.ReadValues("Sheet1!A5:C5");
            OutputValueList("# Origional Values", ori_values);

            // Test1: create and write values to sheet
            IList<IList<object>> write_values = new List<IList<object>> {
                new List<object> { "Test col A", "Test col B", "Test col C" } };
            OutputValueList("# Write Values", write_values);
            IList<IList<object>> response_values = sheet.WriteValues("Sheet1!A5:C5", write_values);
            Assert.IsNotNull(response_values);
            OutputValueList("# Response Values", response_values);
            // loop[ through response values and compare them with the write_values
            for (int row = 0; row < response_values.Count; row++)
                for (int col = 0; col < response_values[row].Count; col++)
                    Assert.AreEqual(write_values[row][col], response_values[row][col]);

            // Test2: reset values to the origional values
            OutputValueList("# Reset Values", sheet.WriteValues("Sheet1!A5:C5", ori_values));
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