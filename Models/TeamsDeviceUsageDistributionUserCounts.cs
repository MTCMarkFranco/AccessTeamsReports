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
    public class TeamsDeviceUsageDistributionUserCounts
    {
        [Name("Report Refresh Date")]
        public DateTime ReportRefreshDate { get; set; }

        [Name("Web")]
        public int Web { get; set; }

        [Name("Windows Phone")]
        public int WindowsPhone { get; set; }

        [Name("Android Phone")]
        public int AndroidPhone { get; set; }

        [Name("iOS")]
        public int iOS { get; set; }

        [Name("Mac")]
        public int Mac { get; set; }

        [Name("Windows")]
        public int Windows { get; set; }

        [Name("Report Period")]
        public int ReportPeriod { get; set; }
       
    }
}
