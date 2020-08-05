using System;
using System.Collections.Generic;
using CloudCare.BL.DataAccess.Services.ReportingService.Reports;

namespace PDFGen_test01
{
    public class CCReportTest
    {

        public static void CreateReport()
        {
            string[] headers =
            {
                "name",
                "group",
            };

           // double[] widths = {30, 28, 23, 18, 18, 25, 25, 33};
            double[] widths = {100,100};

            List<DataModel> list = new List<DataModel>();
             list.Add(new DataModel(){name = "Test1", group = "Test2"}); 
             list.Add(new DataModel(){name = "Test2", group = "Test2"}); 
             list.Add(new DataModel(){name = "Test3", group = "Test2"}); 
             list.Add(new DataModel(){name = "Test4", group = "Test2"}); 
             list.Add(new DataModel(){name = "Test5", group = "Test2"}); 
             list.Add(new DataModel(){name = "Test6", group = "Test2"}); 
             
            var ccReport = new CloudCareReport();
            ccReport.AddDataGrid(list, headers, widths);
            ccReport.CreateDocument("MYTEST.pdf");
            var path = ccReport.DocumentPath;
            Console.WriteLine($"PATH:{path}");
        }

        public class DataModel
        {
           public string name { get; set; }
           public string group { get; set; }
        }
    }
}