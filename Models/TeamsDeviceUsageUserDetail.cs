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
    public class TeamsDeviceUsageUserDetail
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

        [Name("Used Web")]
        public string UsedWeb { get; set; }

        [Name("Used Windows Phone")]
        public string UsedWindowsPhone { get; set; }

        [Name("Used iOS")]
        public string UsediOS { get; set; }

        [Name("Used Mac")]
        public string UsedMac { get; set; }

        [Name("Used Android Phone")]
        public string UsedAndroidPhone { get; set; }

        [Name("Used Windows")]
        public string UsedWindows { get; set; }

        [Name("Report Period")]
        public int ReportPeriod { get; set; }

    }
}
