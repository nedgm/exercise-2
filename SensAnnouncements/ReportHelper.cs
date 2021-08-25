using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace SensAnnouncements
{
    public static class ReportHelper
    {
        public static void FormatOutput(IReadOnlyCollection<SensAnnouncementReportingItem> reportingData, string reportFileName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var report = new ExcelPackage())
            {
                var worksheet = report.Workbook.Worksheets.Add("Announcements");
                AddHeader(worksheet);
                AddBody(worksheet, reportingData);
                worksheet.Cells.AutoFitColumns();
                report.SaveAs(new FileInfo(reportFileName));
            }
        }

        private static void AddHeader(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A1"].Value = "Title";
            worksheet.Cells["B1"].Value = "Url";
            worksheet.Cells["C1"].Value = "SID";
            worksheet.Cells["D1"].Value = "Timestamp";
            worksheet.Cells["A1:D1"].Style.Font.Bold = true;
            worksheet.Cells["A1:D1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells["A1:D1"].Style.Fill.BackgroundColor.SetColor(Color.MediumSeaGreen);
        }

        private static void AddBody(ExcelWorksheet worksheet, IReadOnlyCollection<SensAnnouncementReportingItem> reportingData)
        {
            var reportRow = 2;
            foreach (var reportingItem in reportingData)
            {
                worksheet.Cells[$"A{reportRow}"].Value = reportingItem.Title;
                worksheet.Cells[$"B{reportRow}"].Value = reportingItem.Url;
                worksheet.Cells[$"B{reportRow}"].Style.Font.Color.SetColor(Color.Blue);
                worksheet.Cells[$"B{reportRow}"].Style.Font.UnderLine = true;
                worksheet.Cells[$"B{reportRow}"].SetHyperlink(new Uri(reportingItem.Url));
                worksheet.Cells[$"C{reportRow}"].Value = reportingItem.Sid;
                worksheet.Cells[$"D{reportRow}"].Value = reportingItem.Timestamp;
                
                reportRow++;
            }
        }
    }
}