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
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //replace with you source files
            string infoPath = @"C:\Users\jawad\Desktop\SPORTCRED.xlsx";
            string logosPath = @"C:\Users\jawad\Desktop\Logos.xlsx";

            ExcelFile infoFile = new ExcelFile(infoPath);
            ExcelFile logosFile = new ExcelFile(logosPath);

            List<Game> games = new List<Game>();
            Dictionary<string, string> logos = new Dictionary<string, string>();
            
            ExcelWorksheet logosWorksheet = logosFile.GetPackage().Workbook.Worksheets[0];
            int logosRows = logosWorksheet.Dimension.Rows;

            for (int i = 1; i <= logosRows; i++)
            {
                string team = logosWorksheet.Cells[i, 1].Value.ToString();
                string logo = logosWorksheet.Cells[i, 3].Value.ToString();
                logos.Add(team.Trim(), logo.Trim());
            }

            //Info Excel
            for (int week = 0; week < infoFile.GetPackage().Workbook.Worksheets.Count; week++)
            {
                ExcelWorksheet worksheet = infoFile.GetPackage().Workbook.Worksheets[week];
                int rows = worksheet.Dimension.Rows;
                for (int i = 1; i <= rows; i++)
                {
                    string teamsString = worksheet.Cells[i, 1].Value.ToString();
                    long dateNum = long.Parse(worksheet.Cells[i, 2].Value.ToString());
                    string result = DateTime.FromOADate(dateNum).ToString("yyyy-MM-dd");
                    string winner = worksheet.Cells[i, 3].Value.ToString();
                    string round = "";

                    TeamsNames teams = new TeamsNames(teamsString);
                    teams.SplitTeams();
                    Team teamA = new Team(teams.teamA, logos[teams.teamA]);
                    Team teamB = new Team("", "");

                    if (teams.teamB.Contains("("))
                    {
                        round = teams.teamB.Split('(', ')')[1];
                        teamB.name = teams.teamB.Split('(')[0].Trim();
                    }
                    else
                    {
                        teamB.name = teams.teamB;
                    }

                    teamB.logo = logos[teamB.name.Trim()];
                    Teams teamsObject = new Teams(teamA, teamB);
                    Game game = new Game(result, winner, teamsObject, round);

                    games.Add(game);
                }
            }
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(games.ToArray());
            //replace with you target file
            System.IO.File.WriteAllText(@"C:\Users\jawad\Desktop\tar.json", json);

            Console.WriteLine("Done!");
        }
    }

}
