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
    public class SheetManager
    {
        List<Sheet> sheets;
        SheetService runningService;

        public SheetManager(string cliend_id_location, string application_name)
        {
            runningService = new SheetService();
            runningService.Run(cliend_id_location, application_name);
        }

        public Sheet GetSheet(string spreadsheetId)
        {
            Sheet sheet = Lookup(spreadsheetId);
            if (sheet == null)
            {
                sheet = new Sheet(runningService.Resource(), spreadsheetId);
                sheets.Add(sheet);
            }

            return sheet;
        }

        private Sheet Lookup(string spreadsheetId)
        {
            foreach (Sheet sheet in sheets)
                if (sheet.Id == spreadsheetId)
                    return sheet;

            return null;
        }
    }
}