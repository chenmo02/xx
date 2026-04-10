using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
#pragma warning disable CS0618

namespace WpfApp1.Services
{
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

            using var pkg = new ExcelPackage();

            AddSummarySheet(pkg, result, dbType, tableName);
            AddDetailSheet(pkg, result.Issues);
            AddFieldSummarySheet(pkg, result.Issues, schema);

            var fi = new FileInfo(savePath);
            pkg.SaveAs(fi);
        }

        // ── Sheet1：验证摘要 ──────────────────────────────────
        private static void AddSummarySheet(ExcelPackage pkg, DvValidationResult r, DvDbType dbType, string tableName)
        {
            var ws = pkg.Workbook.Worksheets.Add("验证摘要");
            ws.DefaultColWidth = 24;

            int row = 1;
            void AddRow(string label, string value)
            {
                ws.Cells[row, 1].Value = label;
                ws.Cells[row, 2].Value = value;
                row++;
            }

            AddRow("数据库类型", dbType == DvDbType.SqlServer ? "SQL Server" : "PostgreSQL");
            AddRow("目标表名", tableName);
            AddRow("数据总行数", r.TotalRows.ToString());
            AddRow("已处理行数", r.ProcessedRows.ToString());
            AddRow("错误记录数(按主键)", r.ErrorCount.ToString());
            AddRow("警告记录数(按主键)", r.WarningCount.ToString());
            AddRow("错误明细条数", r.RawErrorCount.ToString());
            AddRow("错误率", r.TotalRows > 0 ? $"{(double)r.ErrorCount / r.TotalRows:P2}" : "0%");
            AddRow("是否完整校验", r.WasCancelled ? "否（中途取消）" : "是");
            AddRow("校验耗时", $"{r.Elapsed.TotalSeconds:F2} 秒");
            AddRow("校验时间", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            // 样式：标签列加粗
            using var labelRange = ws.Cells[1, 1, row - 1, 1];
            labelRange.Style.Font.Bold = true;
            labelRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            labelRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(243, 244, 246));
        }

        // ── Sheet2：错误明细 ──────────────────────────────────
        private static void AddDetailSheet(ExcelPackage pkg, IReadOnlyList<DvIssue> issues)
        {
            var ws = pkg.Workbook.Worksheets.Add("错误明细");
            string[] headers = ["主键", "行号", "源字段", "目标字段", "目标类型", "级别", "错误类型", "实际值", "说明"];
            for (int c = 0; c < headers.Length; c++)
                ws.Cells[1, c + 1].Value = headers[c];

            StyleHeader(ws, 1, headers.Length);

            int row = 2;
            foreach (var issue in issues)
            {
                ws.Cells[row, 1].Value = issue.PrimaryKeyDisplay;
                ws.Cells[row, 2].Value = issue.RowNumber;
                ws.Cells[row, 3].Value = issue.SourceColumnName;
                ws.Cells[row, 4].Value = issue.TargetColumnName;
                ws.Cells[row, 5].Value = issue.TargetDataType;
                ws.Cells[row, 6].Value = issue.LevelText;
                ws.Cells[row, 7].Value = issue.ErrorType;
                ws.Cells[row, 8].Value = issue.ActualValue;
                ws.Cells[row, 9].Value = issue.Message;

                // 错误行红色、警告行橙色
                if (issue.Level == DvValidationLevel.Error)
                    ws.Cells[row, 6].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(220, 38, 38));
                else if (issue.Level == DvValidationLevel.Warning)
                    ws.Cells[row, 6].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(217, 119, 6));
                row++;
            }

            ws.Cells[1, 1, row - 1, headers.Length].AutoFitColumns(8, 60);
        }

        // ── Sheet3：字段汇总 ──────────────────────────────────
        private static void AddFieldSummarySheet(ExcelPackage pkg, IReadOnlyList<DvIssue> issues, IReadOnlyList<DvTargetColumn> schema)
        {
            var ws = pkg.Workbook.Worksheets.Add("字段汇总");
            string[] headers = ["目标字段", "目标类型", "错误数", "警告数", "主要问题类型"];
            for (int c = 0; c < headers.Length; c++)
                ws.Cells[1, c + 1].Value = headers[c];

            StyleHeader(ws, 1, headers.Length);

            var grouped = issues
                .GroupBy(i => i.TargetColumnName)
                .OrderByDescending(g => g.Count(i => i.Level == DvValidationLevel.Error))
                .ToList();

            int row = 2;
            foreach (var g in grouped)
            {
                var col = schema.FirstOrDefault(c => string.Equals(c.ColumnName, g.Key, StringComparison.OrdinalIgnoreCase));
                ws.Cells[row, 1].Value = g.Key;
                ws.Cells[row, 2].Value = col?.DisplayType ?? "";
                ws.Cells[row, 3].Value = g.Count(i => i.Level == DvValidationLevel.Error);
                ws.Cells[row, 4].Value = g.Count(i => i.Level == DvValidationLevel.Warning);
                ws.Cells[row, 5].Value = g.GroupBy(i => i.ErrorType).OrderByDescending(x => x.Count()).FirstOrDefault()?.Key ?? "";
                row++;
            }

            ws.Cells[1, 1, row - 1, headers.Length].AutoFitColumns(8, 50);
        }

        private static void StyleHeader(ExcelWorksheet ws, int row, int cols)
        {
            using var range = ws.Cells[row, 1, row, cols];
            range.Style.Font.Bold = true;
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(78, 110, 242));
            range.Style.Font.Color.SetColor(System.Drawing.Color.White);
        }
    }
}
