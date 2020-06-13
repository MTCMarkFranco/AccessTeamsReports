using CsvHelper.Configuration.Attributes;
using CsvHelper.TypeConversion;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessTeamsReports.Models
{
    public class TeamsUserActivityUserCounts
    {
        [Name("Report Refresh Date")]
        public DateTime ReportRefreshDate { get; set; }

        [Name("Report Date")]
        public DateTime ReportDate { get; set; }

        [Name("Team Chat Messages")]
        public int TeamChatMessages { get; set; }

        [Name("Private Chat Messages")]
        public int PrivateChatMessages { get; set; }

        [Name("Calls")]
        public int Calls { get; set; }

        [Name("Meetings")]
        public int Meetings { get; set; }

        [Name("Report Period")]
        public int ReportPeriod { get; set; }

    }
}
