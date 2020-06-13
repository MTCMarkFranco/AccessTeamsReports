using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Schema;

namespace AccessTeamsReports.Utilities
{
    public static class ReportCustomSerializersOptions
    {
        // User Centric Reports
        public static TeamsUserActivityCountsSerializer teamsUserActivityCountsSerializer = new TeamsUserActivityCountsSerializer();
        public static TeamsUserActivityUserDetailSerializer teamsUserActivityUserDetailSerializer = new TeamsUserActivityUserDetailSerializer();
        public static TeamsUserActivityUserCountsSerializer teamsUserActivityUserCounts = new TeamsUserActivityUserCountsSerializer();

        // Device Centric Reports
        //public static TeamsDeviceUsageUserDetailSerializer teamsDeviceUsageUserDetailSerializer = new TeamsDeviceUsageUserDetailSerializer();
        //public static TeamsDeviceUsageUserCountsSerializer teamsDeviceUsageUserCountsSerializer = new TeamsDeviceUsageUserCountsSerializer();
        //public static TeamsDeviceUsageDistributionUserCountsSerializer teamsDeviceUsageDistributionUserCountsSerializer = new TeamsDeviceUsageDistributionUserCountsSerializer();

    }
}
