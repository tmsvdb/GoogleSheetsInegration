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
    public class Document
    {
        // properties
        string spreadsheetId;
        SpreadsheetsResource resource;
        List<Page> pages;

        //  Class Information
        public string Id { get { return spreadsheetId; }}
        public SpreadsheetsResource lookupResource { get { return resource; } }
        public int numberOfPages { get { return pages.Count; } }

        // Public interface

        public Document(SpreadsheetsResource resource, string spreadsheetId)
        {
            this.resource = resource;
            this.spreadsheetId = spreadsheetId;
            pages = new List<Page>();
        }

        public Page GetPageByTitle (string pageTitle)
        {
            foreach (Page page in pages)
                if (page.peekTitle == pageTitle)
                    return page;

            return null;
        }

        public Page SetPageByTitle(string pageTitle, Page overwritePage)
        {
            for (int i = 0; i < pages.Count; i++)
            {
                pages[i] = overwritePage;
                return pages[i];
            }
            return null;
        }

        // Read spreadsheet values from Googledrive via Google.Apis

        public void ReadSpreadsheet()
        {
            SpreadsheetsResource.GetRequest spreadsheet_request = resource.Get(Id);
            Spreadsheet spreadsheet = spreadsheet_request.Execute();

            if (spreadsheet.Sheets != null)
            {
                foreach (Google.Apis.Sheets.v4.Data.Sheet sheet in spreadsheet.Sheets)
                {
                    string rangeString = new PageRange(
                        sheet.Properties.Title, 0, 0, 
                        (int)sheet.Properties.GridProperties.RowCount, 
                        (int)sheet.Properties.GridProperties.ColumnCount
                    ).ToRangeString();

                    Console.WriteLine("ReadSpreadsheet page using page range: " + rangeString);

                    SpreadsheetsResource.ValuesResource.GetRequest values_request = resource.Values.Get(Id, rangeString);
                    ValueRange response = values_request.Execute();
                    Page new_page = Page.FromObjectList(sheet.Properties.Title, response.Values);
                    pages.Add(new_page);
                }
            }
        }

        // Write spreadsheet values to Googledrive via Google.Apis

        public void WriteSpreadsheet()
        {
            List<Page> newPages = new List<Page>();

            foreach (Page page in pages)
            {
                ValueRange valueRange = new ValueRange();
                valueRange.MajorDimension = "ROWS";//"ROWS";//COLUMNS
                valueRange.Values = Page.ToObjectList(page);

                string range = page.ToPageRange().ToRangeString();
                Console.WriteLine("WriteSpreadsheet page using page range: " + range);

                SpreadsheetsResource.ValuesResource.UpdateRequest request = resource.Values.Update(valueRange, Id, range);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                request.IncludeValuesInResponse = true;

                UpdateValuesResponse response = request.Execute();
                Page new_page = Page.FromObjectList(page.peekTitle, response.UpdatedData.Values);
                newPages.Add(new_page);
            }

            pages = newPages;
        }
    }
}