using System;
using System.Collections.Generic;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetsApiTest
{
    class Program
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "Test";
        static readonly string SpreadsheetId = "1w7bPa_hrH382oVPDwIW9EotY-rzHcj8VHBesYPHPNEg";
        static readonly string sheet = "Student Information";
        static SheetsService service;
        static void Main(string[] args)
        {
            GoogleCredential credential;
            using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }

            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            ReadEntries();
        }
        static void ReadEntries()
        {
            var range = $"{sheet}!B2:B5"; //Sets a range for values to be retrieved from
            var request = service.Spreadsheets.Values.Get(SpreadsheetId, range); //Pulls the specified data from the spreadsheet

            var response = request.Execute();
            var values = response.Values;
            if(values != null && values.Count > 0)
            {
                foreach(var row in values)
                {
                    Console.WriteLine("{0} | {1} | {2} | {3}", row[0], row[1], row[2], row[3]);
                }
                //Creates a for loop with an example of count i where i is the length of the number of rows getting pulled from the spreadsheet, 
                //with the data pulled being placed in an array format(a)

                //row[int x] function takes each value of x position in a 
            }
            else
            {
                Console.WriteLine("No data found.");
            }
        }

        static void CreateEntry()
        {
            var range = $"{sheet}!A:F";
            var valueRange = new ValueRange();

            var objectList = new List<object>(){ "Hello!", "This", "was", "inserted", "via", "C#"};
            valueRange.Values= new List<IList<object>> { objectList };

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendReponse = appendRequest.Execute();
        }

        static void UpdateEntry()
        {
            var range = $"{sheet}!D5";
            var valueRange = new ValueRange();

            var objectList = new List<object>(){ "updated" };
            valueRange.Values = new List<IList<object>> { objectList };

            var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updateResponse = updateRequest.Execute();
        }

        static void DeleteEntry()
        {
            var range = $"{sheet}!A5:F";
            var requestBody = new ClearValuesRequest();

            var deleteRequest = service.Spreadsheets.Values.Clear(requestBody, SpreadsheetId, range);
            var deleteResponse = deleteRequest.Execute();
        }
    }
}

