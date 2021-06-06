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

        static object[] userId = new string[3];

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

            string keyedId;
            Console.WriteLine("Key in someone's User ID to get their information.");
            keyedId = Console.ReadLine();
            ReadEntries();
            Console.WriteLine(userId[0]);
            Console.WriteLine(userId[1]);
            Console.WriteLine(userId[2]);
            // Console.WriteLine(Userid);
        }
        static void ReadEntries()
        {
            var range = $"{sheet}!A2:B4";
            var request = service.Spreadsheets.Values.Get(SpreadsheetId, range);

            var response = request.Execute();
            var values = response.Values;
            if(values != null && values.Count > 0)
            {
                foreach(var row in values)
                {
                    int x = 0;   

                    // Userid = (row[0], row[1]).ToString();
                    // Console.WriteLine(row[0]);
                    // string[] studentids;
                    
                    userId[x] = row[0];
                    Console.WriteLine(x);
                    x++;
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
        }

        static void CreateEntry()
        {
            var range = $"{sheet}!A2:D2";
            var valueRange = new ValueRange();

            var objectList = new List<object>(){ "211530w", "Sun Zizhuo", "1O3", "123456789" };
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

