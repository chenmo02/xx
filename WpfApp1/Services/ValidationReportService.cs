using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
#pragma warning disable CS0618

namespace WpfApp1.Services
{
    /// <summary>
    /// 数据验证报告导出服务。
    /// 负责把一次校验的结果拆成“摘要 / 错误明细 / 字段汇总”三个 Sheet，
    /// 方便业务同学从不同视角查看问题分布。
    /// </summary>
    public static class ValidationReportService
    {
        public static void Generate(
            DvValidationResult result,
            IReadOnlyList<DvTargetColumn> schema,
            IReadOnlyList<DvMappingRow> mappings,
            DvDbType dbType,
            string tableName,
            string savePath)
        {
            ExcelPackage.License.SetNonCommercialPersonal("CCToolbox");

            using var package = new ExcelPackage();

            AddSummarySheet(package, result, dbType, tableName);
            AddDetailSheet(package, result.Issues);
            AddFieldSummarySheet(package, result.Issues, schema);

            var file = new FileInfo(savePath);
            package.SaveAs(file);
        }

        /// <summary>
        /// Sheet1：校验摘要。
        /// 展示总行数、异常记录数、错误项数、耗时等整体指标。
        /// </summary>
        private static void AddSummarySheet(ExcelPackage package, DvValidationResult result, DvDbType dbType, string tableName)
        {
            var worksheet = package.Workbook.Worksheets.Add("校验摘要");
            worksheet.DefaultColWidth = 24;

            int row = 1;
            void AddRow(string label, string value)
            {
                worksheet.Cells[row, 1].Value = label;
                worksheet.Cells[row, 2].Value = value;
                row++;
            }

            AddRow("数据库类型", dbType == DvDbType.SqlServer ? "SQL Server" : "PostgreSQL");
            AddRow("目标表名", tableName);
            AddRow("数据总行数", result.TotalRows.ToString());
            AddRow("已处理行数", result.ProcessedRows.ToString());
            AddRow("异常记录数（按主键/行去重）", result.ErrorCount.ToString());
            AddRow("警告记录数（按主键/行去重）", result.WarningCount.ToString());
            AddRow("错误项数", result.RawErrorCount.ToString());
            AddRow("错误率", result.TotalRows > 0 ? $"{(double)result.ErrorCount / result.TotalRows:P2}" : "0%");
            AddRow("是否完整校验", result.WasCancelled ? "否（中途取消）" : "是");
            AddRow("校验耗时", $"{result.Elapsed.TotalSeconds:F2} 秒");
            AddRow("导出时间", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            using var labelRange = worksheet.Cells[1, 1, row - 1, 1];
            labelRange.Style.Font.Bold = true;
            labelRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            labelRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(243, 244, 246));
        }

        /// <summary>
        /// Sheet2：错误明细。
        /// 一行对应一个错误项或警告项，适合回查具体字段和值。
        /// </summary>
        private static void AddDetailSheet(ExcelPackage package, IReadOnlyList<DvIssue> issues)
        {
            var worksheet = package.Workbook.Worksheets.Add("错误明细");
            string[] headers =
            [
                "主键",
                "行号",
                "源字段",
                "目标字段",
                "目标类型",
                "级别",
                "错误类型",
                "实际值",
                "说明"
            ];

            for (int column = 0; column < headers.Length; column++)
            {
                worksheet.Cells[1, column + 1].Value = headers[column];
            }

            StyleHeader(worksheet, 1, headers.Length);

            int row = 2;
            foreach (var issue in issues)
            {
                worksheet.Cells[row, 1].Value = issue.PrimaryKeyDisplay;
                worksheet.Cells[row, 2].Value = issue.RowNumber;
                worksheet.Cells[row, 3].Value = issue.SourceColumnName;
                worksheet.Cells[row, 4].Value = issue.TargetColumnName;
                worksheet.Cells[row, 5].Value = issue.TargetDataType;
                worksheet.Cells[row, 6].Value = issue.LevelText;
                worksheet.Cells[row, 7].Value = issue.ErrorType;
                worksheet.Cells[row, 8].Value = issue.ActualValue;
                worksheet.Cells[row, 9].Value = issue.Message;

                if (issue.Level == DvValidationLevel.Error)
                {
                    worksheet.Cells[row, 6].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(220, 38, 38));
                }
                else if (issue.Level == DvValidationLevel.Warning)
                {
                    worksheet.Cells[row, 6].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(217, 119, 6));
                }

                row++;
            }

            worksheet.Cells[1, 1, row - 1, headers.Length].AutoFitColumns(8, 60);
        }

        /// <summary>
        /// Sheet3：字段汇总。
        /// 从“字段”维度统计问题数，方便快速定位高频异常字段。
        /// </summary>
        private static void AddFieldSummarySheet(
            ExcelPackage package,
            IReadOnlyList<DvIssue> issues,
            IReadOnlyList<DvTargetColumn> schema)
        {
            var worksheet = package.Workbook.Worksheets.Add("字段汇总");
            string[] headers =
            [
                "目标字段",
                "目标类型",
                "错误数",
                "警告数",
                "主要问题类型"
            ];

            for (int column = 0; column < headers.Length; column++)
            {
                worksheet.Cells[1, column + 1].Value = headers[column];
            }

            StyleHeader(worksheet, 1, headers.Length);

            var groupedIssues = issues
                .GroupBy(issue => issue.TargetColumnName)
                .OrderByDescending(group => group.Count(issue => issue.Level == DvValidationLevel.Error))
                .ToList();

            int row = 2;
            foreach (var group in groupedIssues)
            {
                var targetColumn = schema.FirstOrDefault(column =>
                    string.Equals(column.ColumnName, group.Key, StringComparison.OrdinalIgnoreCase));

                worksheet.Cells[row, 1].Value = group.Key;
                worksheet.Cells[row, 2].Value = targetColumn?.DisplayType ?? "";
                worksheet.Cells[row, 3].Value = group.Count(issue => issue.Level == DvValidationLevel.Error);
                worksheet.Cells[row, 4].Value = group.Count(issue => issue.Level == DvValidationLevel.Warning);
                worksheet.Cells[row, 5].Value = group
                    .GroupBy(issue => issue.ErrorType)
                    .OrderByDescending(issueGroup => issueGroup.Count())
                    .FirstOrDefault()
                    ?.Key ?? "";
                row++;
            }

            worksheet.Cells[1, 1, row - 1, headers.Length].AutoFitColumns(8, 50);
        }

        private static void StyleHeader(ExcelWorksheet worksheet, int row, int columns)
        {
            using var range = worksheet.Cells[row, 1, row, columns];
            range.Style.Font.Bold = true;
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(78, 110, 242));
            range.Style.Font.Color.SetColor(System.Drawing.Color.White);
        }
    }
}
