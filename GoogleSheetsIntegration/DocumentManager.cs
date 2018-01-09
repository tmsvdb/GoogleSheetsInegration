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
    public class DocumentManager
    {
        // Properties
        List<Document> documents;
        APIService runningService;

        // Class Information
        public int numberOfDocuments { get { return documents.Count; } }
        public bool isRunning { get { return runningService != null; } }

        // Public methodes

        public DocumentManager(string cliend_id_location, string application_name)
        {
            runningService = new APIService();
            runningService.Run(cliend_id_location, application_name);
            documents = new List<Document>();
        }

        public Document CreateNewDocument(string documentTitle)
        {
            Document doc = Document.CreateNewDocument(runningService.Resource(), documentTitle);
            if (doc != null)
            {
                doc.ReadSpreadsheet();
                documents.Add(doc);
            }
            return doc;
        }

        public Document LoadDocument(string spreadsheetId)
        {
            Document doc = Lookup(spreadsheetId);
            if (doc == null)
            {
                doc = new Document(runningService.Resource(), spreadsheetId);
                doc.ReadSpreadsheet();
                documents.Add(doc);
            }

            return doc;
        }

        public Document[] GetLoadedDocuments ()
        {
            return documents.ToArray();
        }

        public bool Save(string spreadsheetId)
        {
            Document doc = Lookup(spreadsheetId);
            return Save(doc);
        }
        public bool Save(Document doc)
        {
            return doc.WriteSpreadsheet();
        }

        public bool SaveAll ()
        {
            foreach (Document doc in documents)
                if (!Save(doc)) return false;

            return true;
        }

        // Local implementation

        private Document Lookup(string spreadsheetId)
        {
            foreach (Document doc in documents)
                if (doc.Id == spreadsheetId)
                    return doc;

            return null;
        }
    }
}