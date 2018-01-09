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
        SpreadsheetProperties properties;
        SpreadsheetsResource resource;
        List<Page> pages;

        //  Class Information
        public string Id { get { return spreadsheetId; }}
        public SpreadsheetsResource lookupResource { get { return resource; } }
        public int numberOfPages { get { return pages.Count; } }
        public SpreadsheetProperties peekProperties { get { return properties; } }

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
                if (page.peekProperties.Title == pageTitle)
                    return page;

            return null;
        }

        public void AddPage (string pageTitle)
        {
            var addSheetRequest = new AddSheetRequest();
            addSheetRequest.Properties = new SheetProperties();
            addSheetRequest.Properties.Title = pageTitle;

            AddToBatch(new Request { AddSheet = addSheetRequest });
            Spreadsheet newSheet = BatchUpdate();

            ReadSpreadsheet();
        }

        public void DeletePage(string pageTitle)
        {
            DeletePage(GetPageByTitle(pageTitle));
        }
        public void DeletePage(Page page)
        { 
            var deleteSheetRequest = new DeleteSheetRequest();
            deleteSheetRequest.SheetId = page.peekProperties.SheetId;

            AddToBatch(new Request { DeleteSheet = deleteSheetRequest });
            Spreadsheet newSheet = BatchUpdate();

            ReadSpreadsheet();
        }

        // Read spreadsheet values from Googledrive via Google.Apis

        public bool ReadSpreadsheet()
        {
            SpreadsheetsResource.GetRequest spreadsheet_request = resource.Get(Id);
            Spreadsheet spreadsheet = spreadsheet_request.Execute();

            if (spreadsheet == null)
                return false;

            properties = spreadsheet.Properties;

            if (spreadsheet.Sheets != null)
            {
                pages.Clear();

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

                    Page new_page = Page.FromObjectList(sheet.Properties, response.Values);

                    // TODO: overwriting creates race conditions
                    // because getpage returned page is owned by USER
                    // update a page if it allready exist instead of overwrite
                    //      |
                    //      V
                    pages.Add(new_page);
                }
            }
            return true;
        }

        // Write spreadsheet values to Googledrive via Google.Apis

        public bool WriteSpreadsheet()
        {
            List<Page> newPages = new List<Page>();

            foreach (Page page in pages)
            {
                ValueRange valueRange = new ValueRange();
                valueRange.MajorDimension = "ROWS"; //"ROWS"; //COLUMNS
                valueRange.Values = Page.ToObjectList(page);

                string range = page.ToPageRange().ToRangeString();
                Console.WriteLine("WriteSpreadsheet page using page range: " + range);

                SpreadsheetsResource.ValuesResource.UpdateRequest request = resource.Values.Update(valueRange, Id, range);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                request.IncludeValuesInResponse = true;

                UpdateValuesResponse response = request.Execute();
                Page new_page = Page.FromObjectList(page.peekProperties, response.UpdatedData.Values);
                newPages.Add(new_page);

                if (!CheckPages(page, new_page))
                    return false;
            }

            // TODO: overwriting creates race conditions
            // because getpage returned page is owned by USER
            // update a page if it allready exist instead of overwrite
            //      |
            //      V
            pages = newPages;
            return true;
        }

        // Static Class Features

        public static Document CreateNewDocument(SpreadsheetsResource resource, string title)
        {
            SpreadsheetProperties new_properties = new SpreadsheetProperties();
            new_properties.Title = title;

            return CreateNewDocument(resource, new_properties);
        }
        public static Document CreateNewDocument(SpreadsheetsResource resource, SpreadsheetProperties spreadsheetProperties)
        {
            Spreadsheet new_spreadsheet = new Spreadsheet();
            new_spreadsheet.Properties = spreadsheetProperties;

            return CreateNewDocument(resource, new_spreadsheet);
        }
        public static Document CreateNewDocument(SpreadsheetsResource resource, Spreadsheet newSpreadsheet)
        {
            SpreadsheetsResource.CreateRequest request = resource.Create(newSpreadsheet);
            Spreadsheet result_spreadsheet = request.Execute();

            if (result_spreadsheet != null)
            {
                Document new_doc = new Document(resource, result_spreadsheet.SpreadsheetId);
                return new_doc;
            }
            return null;
        }

        // Local implementation

        private bool CheckPages (Page expected_page, Page test_page)
        {
            // Test basic page info values against each other
            if ((test_page.peekProperties.SheetId != expected_page.peekProperties.SheetId)
                || (test_page.numberOfCells != expected_page.numberOfCells)
                || (test_page.peekHeight != expected_page.peekHeight)
                || (test_page.peekWidth != expected_page.peekWidth))
                return false;

            // Test all page cell values against each other
            foreach (Cell exp_cell in expected_page.peekCells)
                foreach (Cell tst_cell in test_page.peekCells)
                    if (tst_cell.Compare(exp_cell.Id.rowIndex, exp_cell.Id.colIndex))
                        if (tst_cell.Value != exp_cell.Value)
                            return false;

            return true;
        }

        // Manage BatchUpdateSpreadsheetRequest

        BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest;

        private void AddToBatch(Request request)
        {
            if (batchUpdateSpreadsheetRequest == null)
            {
                batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
                batchUpdateSpreadsheetRequest.Requests = new List<Request>();
            }
            batchUpdateSpreadsheetRequest.Requests.Add(request);
        }

        private Spreadsheet BatchUpdate ()
        {
            if (batchUpdateSpreadsheetRequest == null)
                return null;

            SpreadsheetsResource.BatchUpdateRequest request = resource.BatchUpdate(batchUpdateSpreadsheetRequest, spreadsheetId);          
            BatchUpdateSpreadsheetResponse response = request.Execute();
            batchUpdateSpreadsheetRequest.Requests.Clear();

            return response.UpdatedSpreadsheet;
        }
    }
}