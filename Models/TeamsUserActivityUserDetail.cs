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
    public class TeamsUserActivityUserDetail 
    {
        [Name("Report Refresh Date")]
        public DateTime ReportRefreshDate { get; set; }

        [Name("User Principal Name")]
        public string UserPrincipalName { get; set; }

        [Name("Last Activity Date")]
        public DateTime? LastActivityDate { get; set; }

        [Name("Is Deleted")]
        public bool IsDeleted { get; set; }

        [Name("Deleted Date")]
        public DateTime? DeletedDate { get; set; }

        [Name("Assigned Products")]
        public string AssignedProducts { get; set; }

        [Name("Team Chat Message Count")]
        public int TeamChatMessageCount { get; set; }

        [Name("Private Chat Message Count")]
        public int PrivateChatMessageCount { get; set; }

        [Name("Call Count")]
        public int CallCount { get; set; }

        [Name("Meeting Count")]
        public int MeetingCount { get; set; }

        [Name("Has Other Action")]
        public string HasOtherAction { get; set; }

        [Name("Report Period")]
        public int ReportPeriod { get; set; }

    }
}
