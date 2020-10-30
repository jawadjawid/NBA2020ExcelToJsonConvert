using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToJSON
{
    class TeamsNames
    {
        private string teams;
        public string teamA { get; private set; }
        public string teamB { get; private set; }

        public TeamsNames(string teams)
        {
            this.teams = teams;
        }

        public void SplitTeams()
        {
            String[] splittedTeams = teams.Split("vs");
            teamA = splittedTeams[0].Trim();
            teamB = splittedTeams[1];
        }


    }
}
