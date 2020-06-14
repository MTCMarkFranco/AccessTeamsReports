using AccessTeamsReports.Models;
using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.TypeConversion;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using AccessTeamsReports.Utilities;

/// <summary>
///
/// * Device Centric Reports *
/// getTeamsDeviceUsageUserDetail
/// getTeamsDeviceUsageUserCounts
/// getTeamsDeviceUsageDistributionUserCounts
/// 
///  * User Centric Reports *
/// getTeamsUserActivityUserDetail -Done
/// getTeamsUserActivityCounts
/// getTeamsUserActivityUserCounts
/// 
/// </summary>
namespace AccessTeamsReports
{
    #region User Centric Reports
    /// <summary>
    /// Serialization Class for "graphClient.Reports.GetTeamsUserActivityUserDetail()"
    /// </summary>
    public class TeamsUserActivityUserDetailSerializer : Microsoft.Graph.ISerializer
    {

        public T DeserializeObject<T>(string inputString)
        {
            Report report = Helper.DeserializeCSVReportObject<TeamsUserActivityUserDetail>(inputString);
            return (T)Convert.ChangeType(report, typeof(T));
        }

        public T DeserializeObject<T>(Stream stream)
        {
            return default(T);
        }

        public string SerializeObject(object serializeableObject)
        {
              return "[]";
        }

             
    }

    /// <summary>
    /// Serialization Class for "graphClient.Reports.GetTeamsUserActivityCounts()"
    /// </summary>
    public class TeamsUserActivityCountsSerializer : Microsoft.Graph.ISerializer
    {

        public T DeserializeObject<T>(string inputString)
        {
            Report report = Helper.DeserializeCSVReportObject<TeamsUserActivityCounts>(inputString);
            return (T)Convert.ChangeType(report, typeof(T));
        }

        public T DeserializeObject<T>(Stream stream)
        {
            return default(T);
        }

        public string SerializeObject(object serializeableObject)
        {
            return "[]";
        }


    }

    /// <summary>
    /// Serialization Class for "graphClient.Reports.GetTeamsUserActivityUserCounts()"
    /// </summary>
    public class TeamsUserActivityUserCountsSerializer : Microsoft.Graph.ISerializer
    {

        public T DeserializeObject<T>(string inputString)
        {
            Report report = Helper.DeserializeCSVReportObject<TeamsUserActivityUserCounts>(inputString);
            return (T)Convert.ChangeType(report, typeof(T));
        }

        public T DeserializeObject<T>(Stream stream)
        {
            return default(T);
        }

        public string SerializeObject(object serializeableObject)
        {
            return "[]";
        }


    }

    #endregion

    #region Device Centric Reports

    /// <summary>
    /// Serialization Class for "graphClient.Reports.GetTeamsDeviceUsageDistributionUserCounts()"
    /// </summary>
    public class TeamsDeviceUsageDistributionUserCountsSerializer : Microsoft.Graph.ISerializer
    {

        public T DeserializeObject<T>(string inputString)
        {
            Report report = Helper.DeserializeCSVReportObject<TeamsDeviceUsageDistributionUserCounts>(inputString);
            return (T)Convert.ChangeType(report, typeof(T));
        }

        public T DeserializeObject<T>(Stream stream)
        {
            return default(T);
        }

        public string SerializeObject(object serializeableObject)
        {
            return "[]";
        }


    }

    /// <summary>
    /// Serialization Class for "graphClient.Reports.GetTeamsDeviceUsageUserCounts()"
    /// </summary>
    public class TeamsDeviceUsageUserCountsSerializer : Microsoft.Graph.ISerializer
    {

        public T DeserializeObject<T>(string inputString)
        {
            Report report = Helper.DeserializeCSVReportObject<TeamsDeviceUsageUserCounts>(inputString);
            return (T)Convert.ChangeType(report, typeof(T));
        }

        public T DeserializeObject<T>(Stream stream)
        {
            return default(T);
        }

        public string SerializeObject(object serializeableObject)
        {
            return "[]";
        }


    }

    /// <summary>
    /// Serialization Class for "graphClient.Reports.GetTeamsDeviceUsageUserDetail()"
    /// </summary>
    public class TeamsDeviceUsageUserDetailSerializer : Microsoft.Graph.ISerializer
    {

        public T DeserializeObject<T>(string inputString)
        {
            Report report = Helper.DeserializeCSVReportObject<TeamsDeviceUsageUserDetail>(inputString);
            return (T)Convert.ChangeType(report, typeof(T));
        }

        public T DeserializeObject<T>(Stream stream)
        {
            return default(T);
        }

        public string SerializeObject(object serializeableObject)
        {
            return "[]";
        }


    }

    #endregion
}
