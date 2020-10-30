using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToJSON
{
    class Game
    {
        public string date;
        public string winner;
        public Teams teams;
        public string round;

        public Game(string date, string winner, Teams teams, string round)
        {
            this.date = date;
            this.winner = winner;
            this.teams = teams;
            this.round = round;
        }
    }
}
