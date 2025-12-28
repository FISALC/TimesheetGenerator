using System;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.Maui.Storage;

namespace TimesheetGenerator.Maui
{
    public class TimesheetService
    {
        // ====== SETTINGS ======
        private static readonly string OutputDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
        private const string OutputNameFormat = "TimeSheet_{0:0000}_{1:00}_{2}.xlsx";

        // ====== DATA CONSTANTS ======
        private const string ProjectName = "Staff Augmentation";
        private const string EmployeeName = "Mohammed Fisal"; 
        private const string EmployeeRole = ".NET Developer";

        // ====== COLORS ======
        private static readonly XLColor HeaderBg = XLColor.FromHtml("#963634"); // Dark Red/Brown
        private static readonly XLColor TableHeaderBgFix = XLColor.FromHtml("#404040"); // Dark Grey
        private static readonly XLColor WeekendFill = XLColor.FromHtml("#FCE4D6"); // Peach color

        public async Task<string> GenerateTimesheetAsync(int year, int month, string approverName, string approverRole)
        {
            if (!Directory.Exists(OutputDirectory))
                Directory.CreateDirectory(OutputDirectory);

            var timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            var outputFileName = string.Format(CultureInfo.InvariantCulture, OutputNameFormat, year, month, timestamp);
            var outPath = Path.Combine(OutputDirectory, outputFileName);

            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Timesheet");

            // 1. Build the Layout (Headers, Logos, Static info)
            await BuildLayoutAsync(ws, year, month);

            // 2. Fill Data (Days, Weekends, Hours)
            FillTimesheetData(ws, year, month, approverName, approverRole);

            // 3. Print Settings
            ws.PageSetup.PageOrientation = XLPageOrientation.Portrait;
            ws.PageSetup.PaperSize = XLPaperSize.A4Paper;
            ws.PageSetup.FitToPages(1, 1);
            ws.PageSetup.Margins.SetLeft(0.5);
            ws.PageSetup.Margins.SetRight(0.5);
            ws.PageSetup.Margins.SetTop(0.5);
            ws.PageSetup.Margins.SetBottom(0.5);
            ws.PageSetup.CenterHorizontally = true;
            
            wb.SaveAs(outPath);

            return outPath;
        }

        private async Task BuildLayoutAsync(IXLWorksheet ws, int year, int month)
        {
            // Global font settings
            ws.Style.Font.FontName = "Calibri";
            ws.Style.Font.FontSize = 11;

            // == LOGO AREA ==
            try 
            {
                // MAUI Way: Load from Resources/Raw/logo.png
                using var stream = await FileSystem.OpenAppPackageFileAsync("report_logo.png");
                if (stream != null)
                {
                    // XLWorkbook needs a rewindable stream sometimes, but let's try direct
                    // If errors occur, we might need to copy to MemoryStream
                    using var ms = new MemoryStream();
                    await stream.CopyToAsync(ms);
                    ms.Position = 0;

                    var picture = ws.AddPicture(ms)
                                    .MoveTo(ws.Cell("A1"))
                                    .Scale(0.8);
                }
            }
            catch
            {
                ws.Cell("A1").Value = "VISTAS GLOBAL (Logo Missing)";
                ws.Cell("A1").Style.Font.FontSize = 24;
                ws.Cell("A1").Style.Font.FontColor = XLColor.FromHtml("#A64D4D"); // Dark Red
                ws.Cell("A1").Style.Font.Bold = true;
            }

            // == TOP INFO BLOCK (Rows 6-9) ==
            var startDate = new DateTime(year, month, 1);
            var endDate = startDate.AddMonths(1).AddDays(-1);
            string dateFmt = "d-MMM-yy";

            int r = 8; // Moved down from 6 to give Logo more space
            SetupTopInfoRow(ws, r++, "Project Name:", ProjectName);
            SetupTopInfoRow(ws, r++, "Start Date:", startDate.ToString(dateFmt));
            SetupTopInfoRow(ws, r++, "End Date:", endDate.ToString(dateFmt));
            SetupTopInfoRow(ws, r++, "Employee Name", ""); // Blank as per user request 

            // Borders for top block (Rows 8-11, Cols B-H)
            var topBlock = ws.Range(8, 2, 11, 8);
            topBlock.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
            topBlock.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            
            // Fix specific bolding for labels
            ws.Range(8, 2, 11, 2).Style.Font.Bold = true;
        }

        private void SetupTopInfoRow(IXLWorksheet ws, int row, string label, string value)
        {
            ws.Cell(row, 2).Value = label;
            ws.Cell(row, 3).Value = value;
            // Merge value cells C to H
            ws.Range(row, 3, row, 8).Merge();
            ws.Range(row, 3, row, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        }

        private void FillTimesheetData(IXLWorksheet ws, int year, int month, string approverName, string approverRole)
        {
            // Stylistic Red Bar
            int redBarRow = 13; // shifted down due to top block moving
            var redBarRange = ws.Range(redBarRow, 2, redBarRow, 8);
            redBarRange.Merge();
            redBarRange.Style.Fill.BackgroundColor = HeaderBg; // #963634
            redBarRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            
            int headerRow = 14; 
            
            // == HEADERS ==
            string[] headers = { "Day", "Date", "Days\nPresent/Off", "Time-In", "Time-Out", "OverTime\n(Hours)", "Note" };
            int col = 2; // Start at B
            
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = ws.Cell(headerRow, col + i);
                cell.Value = headers[i];
                cell.Style.Fill.BackgroundColor = TableHeaderBgFix;
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Font.Bold = true;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                cell.Style.Alignment.WrapText = true;
            }

            // Adjust column widths
            ws.Column(2).Width = 20; 
            ws.Column(3).Width = 15; 
            ws.Column(4).Width = 12; 
            ws.Column(5).Width = 10; 
            ws.Column(6).Width = 10; 
            ws.Column(7).Width = 12; 
            ws.Column(8).Width = 15; 

            // == DAYS ==
            int startRow = headerRow + 1;
            int daysInMonth = DateTime.DaysInMonth(year, month);
            int currentRow = startRow;

            for (int day = 1; day <= daysInMonth; day++)
            {
                var date = new DateTime(year, month, day);
                bool isWeekend = (date.DayOfWeek == DayOfWeek.Friday || date.DayOfWeek == DayOfWeek.Saturday);

                ws.Cell(currentRow, 2).Value = date.ToString("dddd");
                ws.Cell(currentRow, 3).Value = date.ToString("d-MMM-yyyy");
                
                // Center align everything
                ws.Range(currentRow, 2, currentRow, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                if (isWeekend)
                {
                    ws.Range(currentRow, 2, currentRow, 8).Style.Fill.BackgroundColor = WeekendFill;
                    ws.Cell(currentRow, 4).Value = "0"; // Days Present
                }
                else
                {
                    ws.Cell(currentRow, 4).Value = "1"; // Days Present
                    ws.Cell(currentRow, 5).Value = "7:00";
                    ws.Cell(currentRow, 6).Value = "15:00";
                }

                currentRow++;
            }

            var tableRange = ws.Range(headerRow, 2, currentRow - 1, 8);
            tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thick; 
            
            // == FOOTER SUMMARY ==
            // Summary 1
            ws.Cell(currentRow, 2).Value = "Total working Days Off in this month (Days)";
            
            // Label (B-D)
            var labelRange1 = ws.Range(currentRow, 2, currentRow, 4);
            labelRange1.Merge(); 
            labelRange1.Style.Font.Bold = true;
            labelRange1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            
            // Value (E)
            var valueCell1 = ws.Cell(currentRow, 5);
            valueCell1.Value = 0; 
            valueCell1.Style.Font.Bold = true;
            valueCell1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

            // Empty Right (F-H)
            var emptyRange1 = ws.Range(currentRow, 6, currentRow, 8);
            emptyRange1.Merge();

            // Apply Global Style to the whole row block (B-H)
            var fullRow1 = ws.Range(currentRow, 2, currentRow, 8);
            fullRow1.Style.Fill.BackgroundColor = XLColor.LightGray;
            fullRow1.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            
            labelRange1.Style.Border.RightBorder = XLBorderStyleValues.Thin;
            valueCell1.Style.Border.RightBorder = XLBorderStyleValues.Thin;

            currentRow++;

            // Summary 2
            ws.Cell(currentRow, 2).Value = "Total Overtime (Hours)";
            
            var labelRange2 = ws.Range(currentRow, 2, currentRow, 4);
            labelRange2.Merge();
            labelRange2.Style.Font.Bold = true;
            labelRange2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            
            var valueCell2 = ws.Cell(currentRow, 5);
            valueCell2.Value = "0"; 
            valueCell2.Style.Font.Bold = true;
            valueCell2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

            var emptyRange2 = ws.Range(currentRow, 6, currentRow, 8);
            emptyRange2.Merge();

            var fullRow2 = ws.Range(currentRow, 2, currentRow, 8);
            fullRow2.Style.Fill.BackgroundColor = XLColor.LightGray;
            fullRow2.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            labelRange2.Style.Border.RightBorder = XLBorderStyleValues.Thin;
            valueCell2.Style.Border.RightBorder = XLBorderStyleValues.Thin;
            
            // == SIGNATURES ==
            currentRow += 3; // Gap
            var endMonthDate = new DateTime(year, month, daysInMonth);
            
            // Employee: Blank Name, Blank Role, Prefilled Date
            CreateSignatureBlock(ws, currentRow, "Employee Signature:", "", "", "Name:", "Role:", "Signature:", endMonthDate, includeClassification: false);
            
            currentRow += 8; 
            
            // Approver: Selected Name, Selected Role, Prefilled Date
            CreateSignatureBlock(ws, currentRow, "Approver Signature:", approverName, approverRole, "Name:", "Role:", "Signature:", endMonthDate, includeClassification: true);
        }

        private void CreateSignatureBlock(IXLWorksheet ws, int startRow, string title, string name, string role, string label1, string label2, string label3, DateTime date, bool includeClassification)
        {
            // Title
            ws.Cell(startRow, 2).Value = title;
            ws.Cell(startRow, 2).Style.Font.Bold = true;
            ws.Cell(startRow, 2).Style.Font.Underline = XLFontUnderlineValues.Single;

            int r = startRow + 2;
            
            // Name Row
            ws.Cell(r, 2).Value = label1; 
            ws.Cell(r, 2).Style.Font.Bold = true;
            
            // Dotted line for Name
            var nameValCell = ws.Cell(r, 3);
            nameValCell.Value = name;
            nameValCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Range(r, 3, r, 5).Merge(); // Merge C-E for line
            ws.Range(r, 3, r, 5).Style.Border.BottomBorder = XLBorderStyleValues.Dotted;

            // Date on right
            ws.Cell(r, 7).Value = "Date:";
            ws.Cell(r, 7).Style.Font.Bold = true;
            var dateValCell = ws.Cell(r, 8);
            dateValCell.Value = date.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture); 
            dateValCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(r, 8).Style.Border.BottomBorder = XLBorderStyleValues.Dotted;

            r += 2;
            // Role Row
            ws.Cell(r, 2).Value = label2;
            ws.Cell(r, 2).Style.Font.Bold = true;
            ws.Cell(r, 3).Value = role;
            ws.Range(r, 3, r, 5).Merge();
            ws.Range(r, 3, r, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Range(r, 3, r, 5).Style.Border.BottomBorder = XLBorderStyleValues.Dotted;

            r += 2;
            // Signature Row
            ws.Cell(r, 2).Value = label3;
            ws.Cell(r, 2).Style.Font.Bold = true;
            ws.Range(r, 3, r, 5).Merge();
            ws.Range(r, 3, r, 5).Style.Border.BottomBorder = XLBorderStyleValues.Dotted;

            if (includeClassification)
            {
                r += 2;
                // Classification Row
                ws.Cell(r, 2).Value = "Classification: Internal - داخلي";
                ws.Cell(r, 2).Style.Font.FontColor = XLColor.Blue;
            }
        }
    }
}
