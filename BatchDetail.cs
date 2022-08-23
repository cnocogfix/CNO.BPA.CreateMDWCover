using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace CNO.BPA.CreateMDWCover
{
    public static class BatchDetail
    {
        #region Variables
        public static Dictionary<string, Int32> BatchIssues = null;
        public static string BatchAgeing = string.Empty;
        public static string BatchNo = string.Empty;
        public static string BatchPriority = string.Empty;
        public static string BoxNo = string.Empty;
        public static string DefaultAutoRejectPath = string.Empty;
        public static string DefaultRejectPath = string.Empty;
        public static string DeletedBatchPath = string.Empty;
        public static string Department = string.Empty;
        public static string Direct2Workflow = string.Empty;
        public static string DispatcherProject = string.Empty;
        public static string DispCitrixPath = string.Empty;
        public static string DispServerPath = string.Empty;
        public static string DocumentCount = string.Empty;
        public static string FaxID = string.Empty;
        public static string FaxKey = string.Empty;
        public static string MaxClaimsPerDoc = string.Empty;
        public static string MDWMaxDocCount = string.Empty;
        public static string OriginalBatchNo = string.Empty;
        public static string PrepDate = string.Empty;
        public static string PrepOperator = string.Empty;
        public static string PriorityFactor = string.Empty;
        public static string ReceivedDate = string.Empty;
        public static string ReceivedDateCRD = string.Empty;
        public static string Rescan = string.Empty;
        public static string ScannerID = string.Empty;
        public static string SkipAutoRotate = string.Empty;
        public static string SiteID = string.Empty;
        public static string TrackUser = string.Empty;
        public static string TrackPerformance = string.Empty;
        public static string WorkCategory = string.Empty;  
        #endregion   
        public static Int32 determineBatchPriority()
        {
            Int32 batchPriority = 0;

            DateTime dtReceivedDate = Convert.ToDateTime(ReceivedDate);

            TimeSpan diff = DateTime.Now.Date.Subtract(dtReceivedDate.Date);
            
            int daysPast = Convert.ToInt32(diff.TotalDays);
            int priorityFactor = Convert.ToInt32(PriorityFactor);
            int currentPriority = Convert.ToInt32(BatchPriority);
            batchPriority = currentPriority - (daysPast * priorityFactor);
            //batch priority of zero needs to be caught.
            if (batchPriority <= 0)
            {
                return 1;
            }
            else
            {
                return batchPriority;
            }
        }
        public static void Clear()
        {
            BatchIssues = null;
            BatchAgeing = string.Empty;
            BatchNo = string.Empty;
            BatchPriority = string.Empty;
            BoxNo = string.Empty;
            DefaultAutoRejectPath = string.Empty;
            DefaultRejectPath = string.Empty;
            DeletedBatchPath = string.Empty;
            Department = string.Empty;
            Direct2Workflow = string.Empty;
            DispatcherProject = string.Empty;
            DocumentCount = string.Empty;
            MaxClaimsPerDoc = string.Empty;
            MDWMaxDocCount = string.Empty;
            OriginalBatchNo = string.Empty;
            PrepDate = string.Empty;
            PrepOperator = string.Empty;
            PriorityFactor = string.Empty;
            ReceivedDate = string.Empty;
            ReceivedDateCRD = string.Empty;
            Rescan = string.Empty;
            ScannerID = string.Empty;
            SiteID = string.Empty;
            TrackUser = string.Empty;
            TrackPerformance = string.Empty;
            WorkCategory = string.Empty;
        }
    }
}
