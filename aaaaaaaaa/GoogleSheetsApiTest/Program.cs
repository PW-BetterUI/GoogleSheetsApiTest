using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

        static string resp;

        static int pos = 0;
        static int i = 0;
        static int x = 0;
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

            Console.WriteLine("Key in a User's ID to get their information");
            resp = Console.ReadLine().ToLower();
            ReadEntries(1, resp);
            ReadEntries(2, resp);
        }
        static void ReadEntries(int t, string r)
        {
            var range = $"{sheet}!A2:D5";
            var request = service.Spreadsheets.Values.Get(SpreadsheetId, range);

            var response = request.Execute();
            var values = response.Values;
            if(values != null && values.Count > 0)
            {

                switch(t)
                {
                    case 1:
                    foreach(var row in values)
                    {
                        if(r == row[0].ToString())
                        {
                            Console.WriteLine("User found! Retrieving their information...");
                            pos = i;  
                            return;
                        }
                        else if(i == 3)
                        {
                            Console.WriteLine("User not found! Are you sure you have not entered the wrong information?");
                            Environment.Exit(0);
                            return;
                        }
                        i++;
                    }
                    break;

                    case 2:
                    foreach(var row in values)
                    {
                        if(x == pos)
                        {
                            Console.WriteLine("{0} {1} {2}", row[1], row[2], row[3]);
                            return;
                        }
                        x++;
                    }
                    break;
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

