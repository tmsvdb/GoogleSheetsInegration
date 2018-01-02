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


namespace GoogleSheetsIntegration
{
    
    public class Sheet
    {
        string spreadsheetId;
        SpreadsheetsResource resource;

        public string Id
        {
            get {
                return spreadsheetId;
            }
        }

        public Sheet(SpreadsheetsResource resource, string spreadsheetId)
        {
            this.resource = resource;
            this.spreadsheetId = spreadsheetId;
        }

        public IList<IList<Object>> ReadValues(string range)
        {
            SpreadsheetsResource.ValuesResource.GetRequest request = resource.Values.Get(Id, range);
            ValueRange response = request.Execute();
            return response.Values;
        }

        public IList<IList<Object>> WriteValues(string range)
        {
            ValueRange valueRange = new ValueRange();
            valueRange.MajorDimension = "COLUMNS";//"ROWS";//COLUMNS

            var oblist = new List<object>() { "My Cell Text" };
            valueRange.Values = new List<IList<object>> { oblist };

            SpreadsheetsResource.ValuesResource.UpdateRequest request = resource.Values.Update(valueRange, Id, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            request.IncludeValuesInResponse = true;
            UpdateValuesResponse response = request.Execute();
            return response.UpdatedData.Values;
        }
    }
}