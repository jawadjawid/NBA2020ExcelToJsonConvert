using Nancy.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace ExcelToJSON
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            //replace with you source file
            string path = @"C:\Users\jawad\Desktop\SPORTCRED.xlsx";
            FileInfo fileInfo = new FileInfo(path);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            ExcelPackage package = new ExcelPackage(fileInfo);
            List<Game> games = new List<Game>();

            for (int round = 0; round < package.Workbook.Worksheets.Count; round++)
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[round];
                int rows = worksheet.Dimension.Rows;
                for (int i = 1; i <= rows; i++)
                {
                    string teamsString = worksheet.Cells[i, 1].Value.ToString();
                    long dateNum = long.Parse(worksheet.Cells[i, 2].Value.ToString());
                    string result = DateTime.FromOADate(dateNum).ToString("yyyy-MM-dd");
                    string winner = worksheet.Cells[i, 3].Value.ToString();
                    string output = "";

                    TeamsNames teams = new TeamsNames(teamsString);
                    teams.SplitTeams();
                    Team teamA = new Team(teams.teamA, "");
                    Team teamB = new Team("", "");

                    if (teams.teamB.Contains("("))
                    {
                        output = teams.teamB.Split('(', ')')[1];
                        teamB.name = teams.teamB.Split('(')[0].Trim();
                    }
                    else
                    {
                        teamB.name = teams.teamB;
                    }

                    Teams teamsObject = new Teams(teamA, teamB);
                    Game game = new Game(result, winner, teamsObject, output);

                    games.Add(game);
                }
            }
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(games.ToArray());
            //replace with you target file
            System.IO.File.WriteAllText(@"C:\Users\jawad\Desktop\tar.json", json);

        }
    }

}
