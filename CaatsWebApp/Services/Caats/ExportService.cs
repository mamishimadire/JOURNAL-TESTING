using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using CaatsWebApp.Models.Caats;
using System.Globalization;
using System.IO.Compression;

namespace CaatsWebApp.Services.Caats;

public sealed class ExportService
{
    public ExportResponse Export(CaatsState state, ExportRequest req)
    {
        if (state.MasterRows.Count == 0)
        {
            return new ExportResponse { Success = false, Message = "No run results found. Run tests first." };
        }

        var preferredFolder = ResolveOutputFolder(req.OutputFolder);
        var folder = preferredFolder;
        string? folderWarning = null;

        try
        {
            Directory.CreateDirectory(folder);
        }
        catch (UnauthorizedAccessException)
        {
            folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "JE_Audit");
            Directory.CreateDirectory(folder);
            folderWarning = $"Selected output folder is not writable. Exported to '{folder}' instead.";
        }

        // Include time so each export run creates a distinct file and avoids confusion with same-day names.
        var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var client = string.IsNullOrWhiteSpace(state.Engagement.Client) ? "Client" : state.Engagement.Client.Replace(" ", "_", StringComparison.Ordinal);
        var files = new List<string>();

        if (req.Excel)
        {
            var excelPath = Path.Combine(folder, $"{client}_CAATS_{stamp}.xlsx");
            WriteExcel(state, excelPath);
            files.Add(excelPath);
        }

        if (req.Word)
        {
            var wordPath = Path.Combine(folder, $"{client}_JE_Working_Paper_{stamp}.docx");
            WriteWord(state, wordPath);
            files.Add(wordPath);
        }

        if (req.Csv)
        {
            var masterCsv = Path.Combine(folder, $"{client}_MASTER_{stamp}.csv");
            var reconCsv = Path.Combine(folder, $"{client}_RECON_{stamp}.csv");
            File.WriteAllLines(masterCsv, BuildMasterCsv(state.MasterRows));
            File.WriteAllLines(reconCsv, BuildReconCsv(state.ReconRows));
            files.Add(masterCsv);
            files.Add(reconCsv);
        }

        if (files.Count > 0)
        {
            var zipPath = Path.Combine(folder, $"{client}_CAATS_Reports_{stamp}.zip");
            if (File.Exists(zipPath))
            {
                File.Delete(zipPath);
            }

            using var archive = ZipFile.Open(zipPath, ZipArchiveMode.Create);
            foreach (var file in files)
            {
                if (File.Exists(file))
                {
                    archive.CreateEntryFromFile(file, Path.GetFileName(file));
                }
            }

            files.Add(zipPath);
        }

        return new ExportResponse
        {
            Success = true,
            Message = folderWarning is null
                ? $"Export complete. {files.Count} file(s) created."
                : $"Export complete. {files.Count} file(s) created. {folderWarning}",
            Files = files,
        };
    }

    private static string ResolveOutputFolder(string? rawPath)
    {
        if (string.IsNullOrWhiteSpace(rawPath))
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "JE_Audit");
        }

        var expanded = Environment.ExpandEnvironmentVariables(rawPath.Trim());
        if (expanded.StartsWith("~"))
        {
            var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            expanded = Path.Combine(home, expanded.TrimStart('~', '\\', '/'));
        }

        return Path.GetFullPath(expanded);
    }

    private static void WriteExcel(CaatsState state, string path)
    {
        using var wb = new XLWorkbook();

        bool IsApplicable(string key)
        {
            return !state.LastProcedures.TryGetValue(key, out var enabled) || enabled;
        }

        ProcedureScope? GetScope(string key)
        {
            return state.LastProcedureScopes.TryGetValue(key, out var scope) ? scope : null;
        }

        static bool InScope(MasterRow row, ProcedureScope? scope)
        {
            var effective = scope ?? new ProcedureScope();
            if (!effective.Manual && !effective.SemiAuto && !effective.Auto)
            {
                return false;
            }

            var isManual = row.IsManual == 1;
            var includeManualLike = effective.Manual || effective.SemiAuto;
            var includeAuto = effective.Auto;
            return (isManual && includeManualLike) || (!isManual && includeAuto);
        }

        List<MasterRow> ScopedRows(string key)
        {
            if (!IsApplicable(key))
            {
                return [];
            }

            var scope = GetScope(key);
            return state.MasterRows.Where(r => InScope(r, scope)).ToList();
        }

        List<MasterRow> ScopedFilter(string key, Func<MasterRow, bool> predicate, bool useGroupKey = false)
        {
            var scoped = ScopedRows(key);
            if (scoped.Count == 0)
            {
                return [];
            }

            return FilterByTest(scoped, predicate, useGroupKey);
        }

        List<MasterRow> ScopedDuplicate(string key)
        {
            var scoped = ScopedRows(key);
            return scoped.Where(x => x.TestDuplicate == 1).ToList();
        }

        var wsSummary = wb.Worksheets.Add("Summary");
        wsSummary.Cell(1, 1).Value = "CAATS v15.1 Summary";
        wsSummary.Cell(2, 1).Value = "Client";
        wsSummary.Cell(2, 2).Value = state.Engagement.Client;
        wsSummary.Cell(3, 1).Value = "Total Lines";
        wsSummary.Cell(3, 2).Value = state.MasterRows.Count;
        wsSummary.Cell(4, 1).Value = "Recon Variance";
        wsSummary.Cell(4, 2).Value = state.ReconRows.Count(x => x.Status.Contains("Variance", StringComparison.Ordinal));

        var wsRecon = wb.Worksheets.Add("GL_TB_Recon");
        WriteReconSheet(wsRecon, IsApplicable("completeness") ? state.ReconRows : []);

        WriteMasterSheet(wb.Worksheets.Add("Weekend_Journals"), ScopedFilter("weekend", x => x.TestWeekend == 1));
        WriteMasterSheet(wb.Worksheets.Add("Holiday_Journals"), ScopedFilter("holiday", x => x.TestHoliday == 1));
        WriteMasterSheet(wb.Worksheets.Add("Backdated_Journals"), ScopedFilter("backdated", x => x.TestBackdated == 1));
        WriteMasterSheet(wb.Worksheets.Add("Adjustment_Correction_Reversal"), ScopedFilter("adj_desc", x => x.TestAdjDesc == 1));
        WriteMasterSheet(wb.Worksheets.Add("Journals_Above_Materiality"), ScopedFilter("above_mat", x => x.TestAboveMat == 1));
        WriteMasterSheet(wb.Worksheets.Add("Round_Amount_Analysis"), ScopedFilter("round", x => x.TestRound == 1));
        WriteMasterSheet(wb.Worksheets.Add("Duplicate_Entries"), ScopedDuplicate("duplicate"));
        WriteMasterSheet(wb.Worksheets.Add("Unbalanced_Journals"), ScopedFilter("unbalanced", x => x.TestUnbalanced == 1, useGroupKey: true));
        WriteMasterSheet(wb.Worksheets.Add("Low_FSLI"), ScopedFilter("low_fsli", x => x.TestLowFsli == 1));

        var wsMaster = wb.Worksheets.Add("MASTER_Population");
        WriteMasterSheet(wsMaster, state.MasterRows);

        var wsAccountAnalysis = wb.Worksheets.Add("Account_Analysis");
        WriteAccountAnalysisSheet(wsAccountAnalysis, state.MasterRows);

        var wsUser = wb.Worksheets.Add("User_Analysis");
        WriteUserAnalysisSheet(wsUser, ScopedRows("user"));

        var wsBenford = wb.Worksheets.Add("Benford_Analysis");
        var benfordRows = IsApplicable("benford") ? CaatsAnalysisService.BenfordAnalysis(ScopedRows("benford").Select(x => x.AbsAmount)) : [];
        WriteBenfordSheet(wsBenford, benfordRows);

        var riskRows = state.RiskScoreSummaryRows.Count > 0 ? state.RiskScoreSummaryRows : BuildRiskScoreSummaryRows(state.MasterRows);
        WriteRiskScoreSummarySheet(wb.Worksheets.Add("Risk_Score_Summary"), riskRows);

        var userFullRows = state.UserAnalysisFullRows.Count > 0 ? state.UserAnalysisFullRows : BuildUserAnalysisFullRows(state.MasterRows);
        WriteUserAnalysisFullSheet(wb.Worksheets.Add("User_Analysis_Full"), userFullRows);

        var procedureRows = state.ProcedureRankedRows.Count > 0 ? state.ProcedureRankedRows : BuildProcedureRankedRows(state.MasterRows, state.ReconRows);
        WriteProcedureRankedSheet(wb.Worksheets.Add("Procedure_Results_Ranked"), procedureRows);

        var iriRanked = state.IriRankedRows.Count > 0 ? state.IriRankedRows : BuildIriRankedRows(state.MasterRows, state.Engagement);
        WriteIriRankedSheet(wb.Worksheets.Add("IRI_Ranked_Summary"), iriRanked);

        var iriDetail = state.IriDetailRows.Count > 0 ? state.IriDetailRows : BuildIriDetailRows(state.MasterRows, state.Engagement);
        WriteIriDetailSheet(wb.Worksheets.Add("IRI_Detail"), iriDetail);

        wb.SaveAs(path);
    }

    private static List<MasterRow> FilterByTest(IEnumerable<MasterRow> rows, Func<MasterRow, bool> predicate, bool useGroupKey = false)
    {
        if (useGroupKey)
        {
            var keys = rows.Where(predicate).Select(x => x.JournalKey).Distinct(StringComparer.OrdinalIgnoreCase).ToHashSet(StringComparer.OrdinalIgnoreCase);
            return rows.Where(x => keys.Contains(x.JournalKey)).ToList();
        }

        var composites = rows.Where(predicate).Select(x => x.JournalCompositeKey).Distinct(StringComparer.OrdinalIgnoreCase).ToHashSet(StringComparer.OrdinalIgnoreCase);
        return rows.Where(x => composites.Contains(x.JournalCompositeKey)).ToList();
    }

    private static void WriteMasterSheet(IXLWorksheet ws, List<MasterRow> rows)
    {
        var headers = new[]
        {
            "Account", "Journal Key", "Journal ID", "Posting Date", "CPU Date", "Day Name", "Holiday Name", "Description", "User",
            "Debit", "Credit", "Signed", "Abs", "Manual", "Months Backdated", "Round Category", "Round Pattern",
            "Weekend", "Holiday", "Backdated Journals", "Adjustment/Correction/Reversal", "Journals Above Materiality", "Round Amount Analysis", "Duplicate Entries", "Unbalanced Journals", "Low FSLI",
        };

        for (var i = 0; i < headers.Length; i++)
        {
            ws.Cell(1, i + 1).Value = headers[i];
        }

        var r = 2;
        foreach (var row in rows)
        {
            ws.Cell(r, 1).Value = row.Acct;
            ws.Cell(r, 2).Value = row.JournalKey;
            ws.Cell(r, 3).Value = row.JournalId;
            ws.Cell(r, 4).Value = row.PostingDate;
            ws.Cell(r, 5).Value = row.CpuDate;
            ws.Cell(r, 6).Value = row.DayName;
            ws.Cell(r, 7).Value = row.HolidayName;
            ws.Cell(r, 8).Value = row.Description;
            ws.Cell(r, 9).Value = row.User;
            ws.Cell(r, 10).Value = row.Debit;
            ws.Cell(r, 11).Value = row.Credit;
            ws.Cell(r, 12).Value = row.Signed;
            ws.Cell(r, 13).Value = row.AbsAmount;
            ws.Cell(r, 14).Value = row.ManualReason;
            ws.Cell(r, 15).Value = row.MonthsBackdated;
            ws.Cell(r, 16).Value = row.RoundCategory;
            ws.Cell(r, 17).Value = row.RoundPattern;
            ws.Cell(r, 18).Value = row.TestWeekend;
            ws.Cell(r, 19).Value = row.TestHoliday;
            ws.Cell(r, 20).Value = row.TestBackdated;
            ws.Cell(r, 21).Value = row.TestAdjDesc;
            ws.Cell(r, 22).Value = row.TestAboveMat;
            ws.Cell(r, 23).Value = row.TestRound;
            ws.Cell(r, 24).Value = row.TestDuplicate;
            ws.Cell(r, 25).Value = row.TestUnbalanced;
            ws.Cell(r, 26).Value = row.TestLowFsli;
            r++;
        }

        ws.Columns().AdjustToContents();
    }

    private static void WriteReconSheet(IXLWorksheet ws, List<ReconRow> rows)
    {
        var headers = new[] { "Account Number", "Account Name", "Opening Balance", "Closing Balance", "GL Debit", "GL Credit", "GL Balance", "Difference", "Status", "GL Source" };
        for (var i = 0; i < headers.Length; i++)
        {
            ws.Cell(1, i + 1).Value = headers[i];
        }

        var r = 2;
        foreach (var row in rows)
        {
            ws.Cell(r, 1).Value = row.AccountNumber;
            ws.Cell(r, 2).Value = row.AccountName;
            ws.Cell(r, 3).Value = row.OpeningBalance;
            ws.Cell(r, 4).Value = row.ClosingBalance;
            ws.Cell(r, 5).Value = row.GlDebit;
            ws.Cell(r, 6).Value = row.GlCredit;
            ws.Cell(r, 7).Value = row.GlBalance;
            ws.Cell(r, 8).Value = row.Difference;
            ws.Cell(r, 9).Value = row.Status;
            ws.Cell(r, 10).Value = row.GlSource;
            r++;
        }

        ws.Columns().AdjustToContents();
    }

    private static void WriteBenfordSheet(IXLWorksheet ws, List<BenfordRow> rows)
    {
        ws.Cell(1, 1).Value = "Leading Digit";
        ws.Cell(1, 2).Value = "Count";
        ws.Cell(1, 3).Value = "Actual %";
        ws.Cell(1, 4).Value = "Expected %";
        ws.Cell(1, 5).Value = "Difference %";
        ws.Cell(1, 6).Value = "Status";

        var r = 2;
        foreach (var row in rows)
        {
            ws.Cell(r, 1).Value = row.LeadingDigit;
            ws.Cell(r, 2).Value = row.Count;
            ws.Cell(r, 3).Value = row.ActualPercent;
            ws.Cell(r, 4).Value = row.ExpectedPercent;
            ws.Cell(r, 5).Value = row.DifferencePercent;
            ws.Cell(r, 6).Value = row.Status;
            r++;
        }

        ws.Columns().AdjustToContents();
    }

    private static void WriteUserAnalysisSheet(IXLWorksheet ws, List<MasterRow> rows)
    {
        var headers = new[] { "User", "Type", "Lines", "Total Debit", "Total Credit", "Net (Dr-Cr)", "Total Abs Amt (Net)" };
        for (var i = 0; i < headers.Length; i++)
        {
            ws.Cell(1, i + 1).Value = headers[i];
        }

        var grouped = rows
            .GroupBy(x => new { x.User, x.IsManual })
            .Select(g => new
            {
                User = g.Key.User,
                Type = g.Key.IsManual == 1 ? "Manual" : "Automated",
                Lines = g.Count(),
                TotalDebit = g.Sum(x => x.Debit),
                TotalCredit = g.Sum(x => x.Credit),
            })
            .OrderBy(x => x.User, StringComparer.OrdinalIgnoreCase)
            .ThenBy(x => x.Type, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var r = 2;
        foreach (var row in grouped)
        {
            var net = row.TotalDebit - row.TotalCredit;
            ws.Cell(r, 1).Value = row.User;
            ws.Cell(r, 2).Value = row.Type;
            ws.Cell(r, 3).Value = row.Lines;
            ws.Cell(r, 4).Value = row.TotalDebit;
            ws.Cell(r, 5).Value = row.TotalCredit;
            ws.Cell(r, 6).Value = net;
            ws.Cell(r, 7).Value = Math.Abs(net);
            r++;
        }

        ws.Columns().AdjustToContents();
    }

    private static void WriteAccountAnalysisSheet(IXLWorksheet ws, List<MasterRow> rows)
    {
        ws.Cell(1, 1).Value = "Most Used Accounts";
        var headers = new[] { "Account", "Lines", "Journal Groups", "Total Abs Value" };
        for (var i = 0; i < headers.Length; i++)
        {
            ws.Cell(2, i + 1).Value = headers[i];
        }

        var accountStats = rows
            .GroupBy(x => x.Acct)
            .Select(g => new
            {
                Account = g.Key,
                Lines = g.Count(),
                JournalGroups = g.Select(x => x.JournalKey).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
                TotalAbs = g.Sum(x => x.AbsAmount),
            })
            .ToList();

        var mostUsed = accountStats
            .OrderByDescending(x => x.Lines)
            .ThenByDescending(x => x.TotalAbs)
            .ThenBy(x => x.Account, StringComparer.OrdinalIgnoreCase)
            .Take(20)
            .ToList();

        var leastUsed = accountStats
            .OrderBy(x => x.Lines)
            .ThenBy(x => x.TotalAbs)
            .ThenBy(x => x.Account, StringComparer.OrdinalIgnoreCase)
            .Take(20)
            .ToList();

        var r = 3;
        foreach (var row in mostUsed)
        {
            ws.Cell(r, 1).Value = row.Account;
            ws.Cell(r, 2).Value = row.Lines;
            ws.Cell(r, 3).Value = row.JournalGroups;
            ws.Cell(r, 4).Value = row.TotalAbs;
            r++;
        }

        var leastHeaderRow = r + 2;
        ws.Cell(leastHeaderRow - 1, 1).Value = "Least Used Accounts";
        for (var i = 0; i < headers.Length; i++)
        {
            ws.Cell(leastHeaderRow, i + 1).Value = headers[i];
        }

        r = leastHeaderRow + 1;
        foreach (var row in leastUsed)
        {
            ws.Cell(r, 1).Value = row.Account;
            ws.Cell(r, 2).Value = row.Lines;
            ws.Cell(r, 3).Value = row.JournalGroups;
            ws.Cell(r, 4).Value = row.TotalAbs;
            r++;
        }

        ws.Column(4).Style.NumberFormat.Format = "#,##0.00";
        ws.Columns().AdjustToContents();
    }

    private static void WriteRiskScoreSummarySheet(IXLWorksheet ws, List<RiskScoreSummaryRow> rows)
    {
        ws.Cell(1, 1).Value = "Test";
        ws.Cell(1, 2).Value = "Total Risk Score";
        ws.Cell(1, 3).Value = "Lines Flagged";

        var r = 2;
        foreach (var row in rows)
        {
            ws.Cell(r, 1).Value = row.Test;
            ws.Cell(r, 2).Value = row.TotalRiskScore;
            ws.Cell(r, 3).Value = row.LinesFlagged;
            r++;
        }

        ws.Columns().AdjustToContents();
    }

    private static void WriteUserAnalysisFullSheet(IXLWorksheet ws, List<UserAnalysisFullRow> rows)
    {
        var headers = new[]
        {
            "Full Name", "Manual/Automated", "Line Count Debit", "Total Reporting Amount Debit",
            "Line Count Credit", "Total Reporting Amount Credit", "Total Line Count", "Total Reporting Amount",
        };
        for (var i = 0; i < headers.Length; i++)
        {
            ws.Cell(1, i + 1).Value = headers[i];
        }

        var r = 2;
        foreach (var row in rows)
        {
            ws.Cell(r, 1).Value = row.FullName;
            ws.Cell(r, 2).Value = row.ManualAutomatedDescriptor;
            ws.Cell(r, 3).Value = row.LineCountDebit;
            ws.Cell(r, 4).Value = row.TotalReportingAmountDebit;
            ws.Cell(r, 5).Value = row.LineCountCredit;
            ws.Cell(r, 6).Value = row.TotalReportingAmountCredit;
            ws.Cell(r, 7).Value = row.TotalLineCount;
            ws.Cell(r, 8).Value = row.TotalReportingAmount;
            r++;
        }

        ws.Columns().AdjustToContents();
    }

    private static void WriteProcedureRankedSheet(IXLWorksheet ws, List<ProcedureRankedRow> rows)
    {
        ws.Cell(1, 1).Value = "Rank";
        ws.Cell(1, 2).Value = "Procedure";
        ws.Cell(1, 3).Value = "Count";
        ws.Cell(1, 4).Value = "Ranked Label";

        var r = 2;
        foreach (var row in rows)
        {
            ws.Cell(r, 1).Value = row.Rank;
            ws.Cell(r, 2).Value = row.Procedure;
            ws.Cell(r, 3).Value = row.Count;
            ws.Cell(r, 4).Value = row.RankedLabel;
            r++;
        }

        ws.Columns().AdjustToContents();
    }

    private static void WriteIriRankedSheet(IXLWorksheet ws, List<IriRankedRow> rows)
    {
        ws.Cell(1, 1).Value = "Rank";
        ws.Cell(1, 2).Value = "IRI";
        ws.Cell(1, 3).Value = "Description";
        ws.Cell(1, 4).Value = "Count";
        ws.Cell(1, 5).Value = "Ranked Label";

        var r = 2;
        foreach (var row in rows)
        {
            ws.Cell(r, 1).Value = row.Rank;
            ws.Cell(r, 2).Value = row.Iri;
            ws.Cell(r, 3).Value = row.Description;
            ws.Cell(r, 4).Value = row.Count;
            ws.Cell(r, 5).Value = row.RankedLabel;
            r++;
        }

        ws.Columns().AdjustToContents();
    }

    private static void WriteIriDetailSheet(IXLWorksheet ws, List<IriDetailRow> rows)
    {
        var headers = new[]
        {
            "IRI", "Description", "Account", "Journal Key", "Journal ID", "Posting Date", "CPU Date", "User",
            "Debit", "Credit", "Abs Amount", "Day", "Holiday", "Months Backdated", "Round Category",
        };
        for (var i = 0; i < headers.Length; i++)
        {
            ws.Cell(1, i + 1).Value = headers[i];
        }

        var r = 2;
        foreach (var row in rows)
        {
            ws.Cell(r, 1).Value = row.Iri;
            ws.Cell(r, 2).Value = row.Description;
            ws.Cell(r, 3).Value = row.Acct;
            ws.Cell(r, 4).Value = row.JournalKey;
            ws.Cell(r, 5).Value = row.JournalId;
            ws.Cell(r, 6).Value = row.PostingDate;
            ws.Cell(r, 7).Value = row.CpuDate;
            ws.Cell(r, 8).Value = row.User;
            ws.Cell(r, 9).Value = row.Debit;
            ws.Cell(r, 10).Value = row.Credit;
            ws.Cell(r, 11).Value = row.AbsAmount;
            ws.Cell(r, 12).Value = row.DayName;
            ws.Cell(r, 13).Value = row.HolidayName;
            ws.Cell(r, 14).Value = row.MonthsBackdated;
            ws.Cell(r, 15).Value = row.RoundCategory;
            r++;
        }

        ws.Columns().AdjustToContents();
    }

    private static List<RiskScoreSummaryRow> BuildRiskScoreSummaryRows(List<MasterRow> master)
    {
        var rows = new List<RiskScoreSummaryRow>
        {
            BuildRiskRow("Manual Journals", master.Count(x => x.IsManual == 1), 3.0m),
            BuildRiskRow("Weekend", master.Count(x => x.TestWeekend == 1), 2.0m),
            BuildRiskRow("Holiday", master.Count(x => x.TestHoliday == 1), 2.0m),
            BuildRiskRow("Adjustment/Correction/Reversal Descriptions", master.Count(x => x.TestAdjDesc == 1), 2.0m),
            BuildRiskRow("Backdated Journals", master.Count(x => x.TestBackdated == 1), 3.0m),
            BuildRiskRow("Journals Above Materiality", master.Count(x => x.TestAboveMat == 1), 3.0m),
            BuildRiskRow("Low FSLI", master.Count(x => x.TestLowFsli == 1), 2.0m),
            BuildRiskRow("Round Amount Analysis", master.Count(x => x.TestRound == 1), 2.0m),
            BuildRiskRow("Unbalanced Journals", master.Count(x => x.TestUnbalanced == 1), 3.0m),
            BuildRiskRow("Duplicate Entries", master.Count(x => x.TestDuplicate == 1), 2.0m),
        };

        var total = rows.Sum(x => x.TotalRiskScore);
        rows.Add(new RiskScoreSummaryRow { Test = "TOTAL", TotalRiskScore = decimal.Round(total, 2), LinesFlagged = master.Count });
        return rows;
    }

    private static RiskScoreSummaryRow BuildRiskRow(string test, int linesFlagged, decimal weight)
    {
        return new RiskScoreSummaryRow
        {
            Test = test,
            LinesFlagged = linesFlagged,
            TotalRiskScore = decimal.Round(linesFlagged * weight, 2),
        };
    }

    private static List<UserAnalysisFullRow> BuildUserAnalysisFullRows(List<MasterRow> master)
    {
        return master
            .GroupBy(x => new { x.User, x.IsManual })
            .Select(g => new UserAnalysisFullRow
            {
                FullName = g.Key.User,
                ManualAutomatedDescriptor = g.Key.IsManual == 1 ? "Manual" : "Automated",
                LineCountDebit = g.Count(x => x.Debit > 0),
                TotalReportingAmountDebit = decimal.Round(g.Sum(x => x.Debit), 2),
                LineCountCredit = g.Count(x => x.Credit > 0),
                TotalReportingAmountCredit = decimal.Round(g.Sum(x => x.Credit), 2),
                TotalLineCount = g.Count(),
                TotalReportingAmount = decimal.Round(g.Sum(x => x.AbsAmount), 2),
            })
            .OrderBy(x => x.FullName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(x => x.ManualAutomatedDescriptor, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static List<ProcedureRankedRow> BuildProcedureRankedRows(List<MasterRow> master, List<ReconRow> recon)
    {
        var rows = new List<ProcedureRankedRow>
        {
            new() { Procedure = "Completeness", Count = recon.Count(x => x.Status.Contains("Variance", StringComparison.OrdinalIgnoreCase)) },
            new() { Procedure = "Manual Weekend", Count = master.Count(x => x.TestWeekend == 1) },
            new() { Procedure = "Manual Holiday", Count = master.Count(x => x.TestHoliday == 1) },
            new() { Procedure = "User Analysis", Count = master.Select(x => x.User).Distinct(StringComparer.OrdinalIgnoreCase).Count() },
            new() { Procedure = "Unbalanced Journals", Count = master.Count(x => x.TestUnbalanced == 1) },
            new() { Procedure = "Above Performance Materiality", Count = master.Count(x => x.TestAboveMat == 1) },
            new() { Procedure = "Infrequent Financial Statement Line Item (FSLI)", Count = master.Count(x => x.TestLowFsli == 1) },
            new() { Procedure = "Manual Round Amounts", Count = master.Count(x => x.TestRound == 1) },
            new() { Procedure = "Backdated Entries", Count = master.Count(x => x.TestBackdated == 1) },
            new() { Procedure = "Description Contains Adjustment", Count = master.Count(x => x.TestAdjDesc == 1) },
            new() { Procedure = "Duplicates Entries", Count = master.Count(x => x.TestDuplicate == 1) },
        };

        return rows.OrderByDescending(x => x.Count)
            .ThenBy(x => x.Procedure, StringComparer.OrdinalIgnoreCase)
            .Select((x, idx) =>
            {
                x.Rank = idx + 1;
                x.RankedLabel = $"{idx + 1}: {x.Procedure}";
                return x;
            })
            .ToList();
    }

    private static List<IriRankedRow> BuildIriRankedRows(List<MasterRow> master, EngagementSettings engagement)
    {
        var detail = BuildIriDetailRows(master, engagement);
        var grouped = detail
            .GroupBy(x => new { x.Iri, x.Description })
            .Select(g => new IriRankedRow
            {
                Iri = g.Key.Iri,
                Description = g.Key.Description,
                Count = g.Count(),
            })
            .ToDictionary(x => x.Iri, x => x, StringComparer.OrdinalIgnoreCase);

        var list = IriDescriptions()
            .Select(kv => grouped.TryGetValue(kv.Key, out var row)
                ? row
                : new IriRankedRow { Iri = kv.Key, Description = kv.Value, Count = 0 })
            .OrderByDescending(x => x.Count)
            .ThenBy(x => x.Iri, StringComparer.OrdinalIgnoreCase)
            .Select((x, idx) =>
            {
                x.Rank = idx + 1;
                x.RankedLabel = $"{idx + 1}: {x.Description}";
                return x;
            })
            .ToList();

        return list;
    }

    private static List<IriDetailRow> BuildIriDetailRows(List<MasterRow> master, EngagementSettings engagement)
    {
        var rows = new List<IriDetailRow>(master.Count / 2);
        var fyEnd = engagement.FyEnd?.Date;
        var desc = IriDescriptions();

        foreach (var row in master)
        {
            var isManual = row.IsManual == 1;
            var isBackdated = row.TestBackdated == 1;
            var isReversal = row.Description.Contains("rever", StringComparison.OrdinalIgnoreCase);
            var isRound = row.TestRound == 1;
            var isAbove = row.TestAboveMat == 1;
            var isLowFsli = row.TestLowFsli == 1;

            if (isManual && isBackdated && isReversal)
            {
                rows.Add(ToIriDetail("SNG_IRI_01", desc["SNG_IRI_01"], row));
            }
            if (isManual && isBackdated && isRound && isReversal && isAbove)
            {
                rows.Add(ToIriDetail("SNG_IRI_02", desc["SNG_IRI_02"], row));
            }
            if (isManual && isBackdated && isRound && isAbove)
            {
                rows.Add(ToIriDetail("SNG_IRI_03", desc["SNG_IRI_03"], row));
            }
            if (isManual && isLowFsli && isReversal)
            {
                rows.Add(ToIriDetail("SNG_IRI_04", desc["SNG_IRI_04"], row));
            }
            if (isManual && isLowFsli && isAbove)
            {
                rows.Add(ToIriDetail("SNG_IRI_05", desc["SNG_IRI_05"], row));
            }
            if (fyEnd.HasValue && row.CpuDate.HasValue && row.CpuDate.Value.Date > fyEnd.Value)
            {
                rows.Add(ToIriDetail("SNG_IRI_10", desc["SNG_IRI_10"], row));
            }
        }

        return rows;
    }

    private static IriDetailRow ToIriDetail(string iri, string description, MasterRow row)
    {
        return new IriDetailRow
        {
            Iri = iri,
            Description = description,
            Acct = row.Acct,
            JournalKey = row.JournalKey,
            JournalId = row.JournalId,
            PostingDate = row.PostingDate,
            CpuDate = row.CpuDate,
            User = row.User,
            Debit = row.Debit,
            Credit = row.Credit,
            AbsAmount = row.AbsAmount,
            DayName = row.DayName,
            HolidayName = row.HolidayName,
            MonthsBackdated = row.MonthsBackdated,
            RoundCategory = row.RoundCategory,
        };
    }

    private static Dictionary<string, string> IriDescriptions()
    {
        return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["SNG_IRI_01"] = "Manual Backdated Reversed",
            ["SNG_IRI_02"] = "Manual Backdated Reversed Round Large",
            ["SNG_IRI_03"] = "Manual Backdated Round Large",
            ["SNG_IRI_04"] = "Manual Reversed Infrequent",
            ["SNG_IRI_05"] = "Manual Infrequent Large",
            ["SNG_IRI_06"] = "(Reserved)",
            ["SNG_IRI_07"] = "(Reserved)",
            ["SNG_IRI_08"] = "(Reserved)",
            ["SNG_IRI_09"] = "(Reserved)",
            ["SNG_IRI_10"] = "Journals posted after period close",
        };
    }

    private static void WriteWord(CaatsState state, string path)
    {
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document();
        var body = mainPart.Document.AppendChild(new Body());

        bool IsApplicable(string key)
        {
            return !state.LastProcedures.TryGetValue(key, out var enabled) || enabled;
        }

        ProcedureScope? GetScope(string key)
        {
            return state.LastProcedureScopes.TryGetValue(key, out var scope) ? scope : null;
        }

        static bool InScope(MasterRow row, ProcedureScope? scope)
        {
            var effective = scope ?? new ProcedureScope();
            if (!effective.Manual && !effective.SemiAuto && !effective.Auto)
            {
                return false;
            }

            var isManual = row.IsManual == 1;
            var includeManualLike = effective.Manual || effective.SemiAuto;
            var includeAuto = effective.Auto;
            return (isManual && includeManualLike) || (!isManual && includeAuto);
        }

        List<MasterRow> ScopedRows(string key)
        {
            if (!IsApplicable(key))
            {
                return [];
            }

            var scope = GetScope(key);
            return state.MasterRows.Where(r => InScope(r, scope)).ToList();
        }

        List<MasterRow> ScopedFilter(string key, Func<MasterRow, bool> predicate, bool useGroupKey = false)
        {
            var scoped = ScopedRows(key);
            return scoped.Count == 0 ? [] : FilterByTest(scoped, predicate, useGroupKey);
        }

        var reportDate = state.Engagement.SigningDate?.ToString("yyyy-MM-dd") ?? DateTime.Now.ToString("yyyy-MM-dd");
        var glDr = state.MasterRows.Sum(x => x.Debit);
        var glCr = state.MasterRows.Sum(x => x.Credit);
        var glDiff = Math.Abs(glDr - glCr);
        var reconAgree = state.ReconRows.Count(x => x.Status.Contains("Agrees", StringComparison.Ordinal));
        var reconVariance = state.ReconRows.Count(x => x.Status.Contains("Variance", StringComparison.Ordinal));

        var weekendInfoRows = ScopedFilter("weekend", x => x.TestWeekendInfo == 1);
        var weekendExceptionRows = ScopedFilter("weekend", x => x.TestWeekend == 1);
        var holidayInfoRows = ScopedFilter("holiday", x => x.TestHolidayInfo == 1);
        var holidayExceptionRows = ScopedFilter("holiday", x => x.TestHoliday == 1);
        var backdatedRows = ScopedFilter("backdated", x => x.TestBackdated == 1);
        var adjRows = ScopedFilter("adj_desc", x => x.TestAdjDesc == 1);
        var aboveMatRows = ScopedFilter("above_mat", x => x.TestAboveMat == 1);
        var roundRows = ScopedFilter("round", x => x.TestRound == 1);
        var duplicateRows = ScopedRows("duplicate").Where(x => x.TestDuplicate == 1).ToList();
        var unbalancedRows = ScopedFilter("unbalanced", x => x.TestUnbalanced == 1, useGroupKey: true);
        var lowFsliRows = ScopedFilter("low_fsli", x => x.TestLowFsli == 1);

        AppendTitleBlock(body, state, reportDate);
        AddTableOfContents(body);

        body.Append(Heading("1. Engagement Information", 2));
        AddSimpleTable(body,
            ["Information Required", "Input"],
            [
                ["Parent Name", state.Engagement.Parent],
                ["Client Name", state.Engagement.Client],
                ["Engagement Name", state.Engagement.Engagement],
                ["General Ledger System", state.Engagement.GlSystem],
                ["Financial Year Start", state.Engagement.FyStart?.ToString("yyyy-MM-dd") ?? string.Empty],
                ["Financial Year End", state.Engagement.FyEnd?.ToString("yyyy-MM-dd") ?? string.Empty],
                ["Period Reviewed Start", state.Engagement.PeriodStart?.ToString("yyyy-MM-dd") ?? string.Empty],
                ["Period Reviewed End", state.Engagement.PeriodEnd?.ToString("yyyy-MM-dd") ?? string.Empty],
                ["Report Signing Date", reportDate],
                ["Materiality", $"R{state.Engagement.Materiality:N2}"],
                ["Performance Materiality", $"R{state.Engagement.PerformanceMateriality:N2}"],
                ["Prepared By", state.Engagement.Auditor],
                ["Manager", state.Engagement.Manager],
            ]);
        AddAuditorCommentBlock(body, "Engagement setup completeness and appropriateness.");

        AddJournalCaatObjectiveSection(body, state);
        AddAnalysisStepsSection(body, state);

        body.Append(Heading("4. Executive Summary", 2));
        AddSimpleTable(body,
            ["Metric", "Result"],
            [
                ["Total GL Lines", state.MasterRows.Count.ToString("N0")],
                ["Total Debit", $"R{glDr:N2}"],
                ["Total Credit", $"R{glCr:N2}"],
                ["Debit/Credit Status", glDiff < 1m ? "Balanced" : $"Difference R{glDiff:N2}"],
                ["Recon Accounts", state.ReconRows.Count.ToString("N0")],
                ["Recon Agrees", reconAgree.ToString("N0")],
                ["Recon Variances", reconVariance.ToString("N0")],
                ["Manual Lines", state.MasterRows.Count(x => x.IsManual == 1).ToString("N0")],
                ["Benford Flagged Digits", state.BenfordRows.Count(x => x.Status.Contains("⚠", StringComparison.Ordinal)).ToString("N0")],
            ]);

        body.Append(Heading("5. Procedure Result Overview", 2));
        AddSimpleTable(body,
            ["Procedure", "Status", "Journal Groups", "Lines", "Detailed Excel Sheet"],
            [
                BuildOverviewRow("Journal Completeness (GL/TB Recon)", IsApplicable("completeness"), state.ReconRows.Count, state.ReconRows.Count, "GL_TB_Recon"),
                BuildOverviewRow("Weekend Manual Journals", IsApplicable("weekend"), DistinctJournalCount(state.Engagement.WeekendNormal ? weekendInfoRows : weekendExceptionRows), (state.Engagement.WeekendNormal ? weekendInfoRows : weekendExceptionRows).Count, "Weekend_Journals"),
                BuildOverviewRow("Public Holiday Manual Journals", IsApplicable("holiday"), DistinctJournalCount(state.Engagement.HolidayNormal ? holidayInfoRows : holidayExceptionRows), (state.Engagement.HolidayNormal ? holidayInfoRows : holidayExceptionRows).Count, "Holiday_Journals"),
                BuildOverviewRow("Backdated Journal Entries", IsApplicable("backdated"), DistinctJournalCount(backdatedRows), backdatedRows.Count, "Backdated"),
                BuildOverviewRow("Adj/Correc Descriptions", IsApplicable("adj_desc"), DistinctJournalCount(adjRows), adjRows.Count, "Adj_Desc"),
                BuildOverviewRow("Journals Above Performance Materiality", IsApplicable("above_mat"), DistinctJournalCount(aboveMatRows), aboveMatRows.Count, "Above_Materiality"),
                BuildOverviewRow("Account Analysis", true, DistinctJournalCount(state.MasterRows), state.MasterRows.Count, "MASTER_Population"),
                BuildOverviewRow("Round Amount Analysis", IsApplicable("round"), DistinctJournalCount(roundRows), roundRows.Count, "Round_Amounts"),
                BuildOverviewRow("Duplicate Entries", IsApplicable("duplicate"), DistinctJournalCount(duplicateRows), duplicateRows.Count, "Duplicates"),
                BuildOverviewRow("Unbalanced Journals", IsApplicable("unbalanced"), DistinctJournalCount(unbalancedRows, useGroupKey: true), unbalancedRows.Count, "Unbalanced"),
                BuildOverviewRow("Low FSLI", IsApplicable("low_fsli"), DistinctJournalCount(lowFsliRows), lowFsliRows.Count, "Low_FSLI"),
                BuildOverviewRow("User Analysis", IsApplicable("user"), state.MasterRows.Select(x => x.User).Distinct(StringComparer.OrdinalIgnoreCase).Count(), state.MasterRows.Count, "User_Analysis"),
                BuildOverviewRow("Benford's Law", IsApplicable("benford"), state.BenfordRows.Count(x => x.Status.Contains("⚠", StringComparison.Ordinal)), state.MasterRows.Count, "Benford_Analysis"),
            ]);
        AddAuditorCommentBlock(body, "Overall conclusion and significant exceptions noted from the above procedures.");

        body.Append(Heading("6.1 Journal Completeness", 2));
        body.Append(Paragraph("Objective: Confirm GL account balances reconcile to TB closing balances."));
        body.Append(Paragraph("Method: Reconciliation uses gl_raw account balances and compares TB Closing Balance to GL Balance per account."));
        AddSimpleTable(body,
            ["Metric", "Result"],
            [
                ["Accounts Reconciled", state.ReconRows.Count.ToString("N0")],
                ["Accounts Agreeing", reconAgree.ToString("N0")],
                ["Accounts with Variance", reconVariance.ToString("N0")],
                ["Largest Absolute Variance", $"R{state.ReconRows.Select(r => Math.Abs(r.Difference)).DefaultIfEmpty(0m).Max():N2}"],
            ]);
        AddSimpleTable(body,
            ["Top Variance Account", "Account Name", "Difference", "Status", "GL Source"],
            state.ReconRows
                .OrderByDescending(r => Math.Abs(r.Difference))
                .Take(10)
                .Select(r => new[]
                {
                    r.AccountNumber,
                    r.AccountName,
                    $"R{r.Difference:N2}",
                    r.Status,
                    r.GlSource,
                }),
            10);
        body.Append(Paragraph("Detailed account-level reconciliation and full populations are available in Excel sheet: GL_TB_Recon."));
        AddAuditorCommentBlock(body, "Completeness conclusion, root causes for variances, and management follow-up.");

        var weekendBaseRows = state.Engagement.WeekendNormal ? weekendInfoRows : weekendExceptionRows;
        body.Append(Heading("6.2 Weekend Manual Journals", 2));
        body.Append(Paragraph("Objective: Identify manual journals posted on weekends."));
        body.Append(Paragraph("Method: Weekend check based on mapped weekend date column; day name is derived directly from posting/check date."));
        AddSimpleTable(body,
            ["Metric", "Result"],
            [
                ["Industry Setting", state.Engagement.WeekendNormal ? "Weekend journals treated as informational" : "Weekend journals treated as exceptions"],
                ["Weekend Journal Groups", DistinctJournalCount(weekendBaseRows).ToString("N0")],
                ["Weekend Lines", weekendBaseRows.Count.ToString("N0")],
            ]);
        AddSimpleTable(body,
            ["Day", "Manual Lines", "Journal Groups"],
            BuildDayBreakdownRows(weekendInfoRows),
            20);
        body.Append(Paragraph("Detailed weekend populations are available in Excel sheet: Weekend_Journals."));
        AddAuditorCommentBlock(body, "Weekend journal rationale and whether timing is expected for the business.");

        var holidayBaseRows = state.Engagement.HolidayNormal ? holidayInfoRows : holidayExceptionRows;
        body.Append(Heading("6.3 Public Holiday Manual Journals (SA Calendar)", 2));
        body.Append(Paragraph("Objective: Identify manual journals posted on South African public holidays."));
        body.Append(Paragraph("Method: SA holiday names are mapped from the built-in South African public holiday calendar (including Good Friday and observed dates)."));
        AddSimpleTable(body,
            ["Metric", "Result"],
            [
                ["Industry Setting", state.Engagement.HolidayNormal ? "Holiday journals treated as informational" : "Holiday journals treated as exceptions"],
                ["Holiday Journal Groups", DistinctJournalCount(holidayBaseRows).ToString("N0")],
                ["Holiday Lines", holidayBaseRows.Count.ToString("N0")],
            ]);
        AddSimpleTable(body,
            ["Holiday Name", "Day", "Manual Lines", "Journal Groups"],
            BuildHolidayBreakdownRows(holidayInfoRows),
            30);
        body.Append(Paragraph("Detailed holiday populations are available in Excel sheet: Holiday_Journals."));
        AddAuditorCommentBlock(body, "Holiday posting justification and whether postings were authorized and time-critical.");

        WriteTestSection(
            body,
            "6.4 Backdated Journal Entries",
            "Identify journals where Creation/CPU Year-Month is later than Posting Year-Month.",
            "Backdated",
            backdatedRows,
            IsApplicable("backdated"));

        WriteTestSection(
            body,
            "6.5 Adj/Correc Description Journals",
            "Identify descriptions containing adjustment/correction/reversal keywords.",
            "Adj_Desc",
            adjRows,
            IsApplicable("adj_desc"));

        WriteTestSection(
            body,
            "6.6 Journals Above Performance Materiality",
            $"Identify journals with absolute value >= R{state.Engagement.PerformanceMateriality:N0}.",
            "Above_Materiality",
            aboveMatRows,
            IsApplicable("above_mat"));

        WriteTestSection(
            body,
            "6.7 Round Amount Analysis",
            "Identify v14 round amount patterns (10,000 multiples and 9,999 pattern).",
            "Round_Amounts",
            roundRows,
            IsApplicable("round"));

        WriteTestSection(
            body,
            "6.8 Duplicate Journal Entries",
            "Identify exact duplicate lines using key/date/amount signature.",
            "Duplicates",
            duplicateRows,
            IsApplicable("duplicate"));

        WriteTestSection(
            body,
            "6.9 Unbalanced Journal Entries",
            "Identify journals where grouped debit and credit totals differ by at least R1.",
            "Unbalanced",
            unbalancedRows,
            IsApplicable("unbalanced"));

        WriteTestSection(
            body,
            "6.10 Low FSLI Accounts",
            $"Identify account-period pairs with fewer than {state.Engagement.LowFsliThreshold} postings.",
            "Low_FSLI",
            lowFsliRows,
            IsApplicable("low_fsli"));

        body.Append(Heading("6.11 Account Analysis", 2));
        body.Append(Paragraph("Objective: Identify most and least used accounts throughout the period under review."));

        var accountAnalysis = state.MasterRows
            .GroupBy(x => x.Acct)
            .Select(g => new
            {
                Account = g.Key,
                Lines = g.Count(),
                JournalGroups = g.Select(x => x.JournalKey).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
                TotalAbs = g.Sum(x => x.AbsAmount),
            })
            .OrderByDescending(x => x.Lines)
            .ThenBy(x => x.Account, StringComparer.OrdinalIgnoreCase)
            .ToList();

        AddSimpleTable(body,
            ["Most Used Account", "Lines", "Journal Groups", "Total Abs Value"],
            accountAnalysis.Take(10).Select(x => new[]
            {
                x.Account,
                x.Lines.ToString("N0", CultureInfo.InvariantCulture),
                x.JournalGroups.ToString("N0", CultureInfo.InvariantCulture),
                $"R{x.TotalAbs:N2}",
            }),
            10);

        AddSimpleTable(body,
            ["Least Used Account", "Lines", "Journal Groups", "Total Abs Value"],
            accountAnalysis
                .OrderBy(x => x.Lines)
                .ThenBy(x => x.TotalAbs)
                .ThenBy(x => x.Account, StringComparer.OrdinalIgnoreCase)
                .Take(10)
                .Select(x => new[]
                {
                    x.Account,
                    x.Lines.ToString("N0", CultureInfo.InvariantCulture),
                    x.JournalGroups.ToString("N0", CultureInfo.InvariantCulture),
                    $"R{x.TotalAbs:N2}",
                }),
            10);

        body.Append(Paragraph("Detailed account summaries are available in Excel sheet: Account_Analysis. Full line-level details remain in MASTER_Population."));
        AddAuditorCommentBlock(body, "Interpretation of most/least used accounts and potential risk concentration areas.");

        body.Append(Heading("6.12 User Analysis", 2));
        body.Append(Paragraph("Objective: Summarize activity by user and manual/automated profile."));
        var userAnalysis = ScopedRows("user")
            .GroupBy(x => new { x.User, x.IsManual })
            .Select(g => new
            {
                g.Key.User,
                Type = g.Key.IsManual == 1 ? "Manual" : "Automated",
                Lines = g.Count(),
                Debit = g.Sum(x => x.Debit),
                Credit = g.Sum(x => x.Credit),
            })
            .OrderByDescending(x => Math.Abs(x.Debit - x.Credit))
            .ThenBy(x => x.User, StringComparer.OrdinalIgnoreCase)
            .ToList();
        AddSimpleTable(body,
            ["User", "Type", "Lines", "Total Debit", "Total Credit", "Net (Dr-Cr)", "Total Abs Amt (Net)"],
            userAnalysis.Take(20).Select(x =>
            {
                var net = x.Debit - x.Credit;
                return new[]
                {
                    x.User,
                    x.Type,
                    x.Lines.ToString("N0"),
                    $"R{x.Debit:N2}",
                    $"R{x.Credit:N2}",
                    $"R{net:N2}",
                    $"R{Math.Abs(net):N2}",
                };
            }));
        body.Append(Paragraph($"Showing top {Math.Min(20, userAnalysis.Count):N0} user/type combinations by absolute net movement."));
        body.Append(Paragraph("Detailed user-level lines are available in Excel sheet: User_Analysis."));
        AddAuditorCommentBlock(body, "Users requiring further investigation and reason for elevated activity.");

        body.Append(Heading("6.13 Benford's Law Analysis", 2));
        body.Append(Paragraph("Objective: Compare observed leading-digit distribution to Benford expected distribution."));
        body.Append(Paragraph("Method: Digits with deviation greater than 5 percentage points are flagged."));
        AddSimpleTable(body,
            ["Leading Digit", "Count", "Actual %", "Expected %", "Difference %", "Status"],
            (IsApplicable("benford") ? CaatsAnalysisService.BenfordAnalysis(ScopedRows("benford").Select(x => x.AbsAmount)) : new List<BenfordRow>())
            .Select(x => new[]
            {
                x.LeadingDigit.ToString(CultureInfo.InvariantCulture),
                x.Count.ToString("N0"),
                $"{x.ActualPercent:N1}%",
                $"{x.ExpectedPercent:N1}%",
                $"{x.DifferencePercent:N1}%",
                x.Status,
            }));
        body.Append(Paragraph("Detailed Benford inputs and values are available in Excel sheet: Benford_Analysis."));
        AddAuditorCommentBlock(body, "Assessment of anomalous leading-digit patterns and planned follow-up procedures.");

        var riskSummaryRows = state.RiskScoreSummaryRows.Count > 0 ? state.RiskScoreSummaryRows : BuildRiskScoreSummaryRows(state.MasterRows);
        var userFullRows = state.UserAnalysisFullRows.Count > 0 ? state.UserAnalysisFullRows : BuildUserAnalysisFullRows(state.MasterRows);
        var procedureRankedRows = state.ProcedureRankedRows.Count > 0 ? state.ProcedureRankedRows : BuildProcedureRankedRows(state.MasterRows, state.ReconRows);
        var iriRankedRows = state.IriRankedRows.Count > 0 ? state.IriRankedRows : BuildIriRankedRows(state.MasterRows, state.Engagement);
        var iriDetailRows = state.IriDetailRows.Count > 0 ? state.IriDetailRows : BuildIriDetailRows(state.MasterRows, state.Engagement);

        body.Append(Heading("6.14 Total Risk Score Summary", 2));
        AddSimpleTable(body,
            ["Test", "Total Risk Score", "Lines Flagged"],
            riskSummaryRows.Select(x => new[]
            {
                x.Test,
                x.TotalRiskScore.ToString("N2", CultureInfo.InvariantCulture),
                x.LinesFlagged.ToString("N0", CultureInfo.InvariantCulture),
            }));
        body.Append(Paragraph("Detailed risk-score calculations are available in Excel sheet: Risk_Score_Summary."));
        AddAuditorCommentBlock(body, "Review of cumulative risk profile and prioritization of follow-up work.");

        body.Append(Heading("6.15 Full Z_USER_ANALYSIS Breakdown", 2));
        AddSimpleTable(body,
            ["Full Name", "Manual/Automated", "Line Count Debit", "Total Amt Debit", "Line Count Credit", "Total Amt Credit", "Total Line Count", "Total Reporting Amount"],
            userFullRows.Take(200).Select(x => new[]
            {
                x.FullName,
                x.ManualAutomatedDescriptor,
                x.LineCountDebit.ToString("N0", CultureInfo.InvariantCulture),
                $"R{x.TotalReportingAmountDebit:N2}",
                x.LineCountCredit.ToString("N0", CultureInfo.InvariantCulture),
                $"R{x.TotalReportingAmountCredit:N2}",
                x.TotalLineCount.ToString("N0", CultureInfo.InvariantCulture),
                $"R{x.TotalReportingAmount:N2}",
            }),
            200);
        body.Append(Paragraph("Detailed full population is available in Excel sheet: User_Analysis_Full."));
        AddAuditorCommentBlock(body, "Assessment of unusual posting behavior per user and manual/automated profile.");

        body.Append(Heading("6.16 Procedure Results Ranked", 2));
        AddSimpleTable(body,
            ["Rank", "Procedure", "Count", "Ranked Label"],
            procedureRankedRows.Select(x => new[]
            {
                x.Rank.ToString(CultureInfo.InvariantCulture),
                x.Procedure,
                x.Count.ToString("N0", CultureInfo.InvariantCulture),
                x.RankedLabel,
            }));
        body.Append(Paragraph("Detailed ranked output is available in Excel sheet: Procedure_Results_Ranked."));
        AddAuditorCommentBlock(body, "Reasonableness of ranking and alignment with audit risk expectations.");

        body.Append(Heading("6.17 IRI Ranked Summary and Triggered Detail", 2));
        AddSimpleTable(body,
            ["Rank", "IRI", "Description", "Count", "Ranked Label"],
            iriRankedRows.Select(x => new[]
            {
                x.Rank.ToString(CultureInfo.InvariantCulture),
                x.Iri,
                x.Description,
                x.Count.ToString("N0", CultureInfo.InvariantCulture),
                x.RankedLabel,
            }));

        foreach (var iri in iriRankedRows.Where(x => x.Count > 0))
        {
            var detailForIri = iriDetailRows.Where(x => string.Equals(x.Iri, iri.Iri, StringComparison.OrdinalIgnoreCase)).ToList();
            body.Append(Paragraph($"IRI Detail: {iri.Iri} - {iri.Description} ({iri.Count:N0} hit(s))"));
            AddSimpleTable(body,
                ["Account", "Journal Key", "Doc ID", "Posting Date", "CPU Date", "User", "Debit", "Credit", "Abs", "Day", "Holiday", "Months Backdated", "Round"],
                detailForIri.Take(50).Select(x => new[]
                {
                    x.Acct,
                    x.JournalKey,
                    x.JournalId,
                    x.PostingDate?.ToString("yyyy-MM-dd") ?? string.Empty,
                    x.CpuDate?.ToString("yyyy-MM-dd") ?? string.Empty,
                    x.User,
                    $"R{x.Debit:N2}",
                    $"R{x.Credit:N2}",
                    $"R{x.AbsAmount:N2}",
                    x.DayName,
                    x.HolidayName,
                    x.MonthsBackdated.ToString(CultureInfo.InvariantCulture),
                    x.RoundCategory,
                }),
                50);
        }
        body.Append(Paragraph("Detailed IRI outputs are available in Excel sheets: IRI_Ranked_Summary and IRI_Detail."));
        AddAuditorCommentBlock(body, "Interpretation of IRI triggers and conclusion on potential management override risk.");

        body.Append(Heading("7. Annexure A - Excel Export Index", 2));
        AddSimpleTable(body,
            ["#", "Description", "Sheet"],
            [
                ["1", "GL/TB Reconciliation", "GL_TB_Recon"],
                ["2", "Weekend Journals", "Weekend_Journals"],
                ["3", "Holiday Journals", "Holiday_Journals"],
                ["4", "Backdated Entries", "Backdated"],
                ["5", "Adjustment Descriptions", "Adj_Desc"],
                ["6", "Above Materiality", "Above_Materiality"],
                ["7", "Round Amounts", "Round_Amounts"],
                ["8", "Duplicate Entries", "Duplicates"],
                ["9", "Unbalanced Journals", "Unbalanced"],
                ["10", "Low FSLI Accounts", "Low_FSLI"],
                ["11", "User Analysis", "User_Analysis"],
                ["12", "Benford Analysis", "Benford_Analysis"],
                ["13", "Full Master Population", "MASTER_Population"],
                ["14", "Account Analysis", "Account_Analysis"],
                ["15", "Risk Score Summary", "Risk_Score_Summary"],
                ["16", "Full Z_USER_ANALYSIS", "User_Analysis_Full"],
                ["17", "Procedure Ranked Results", "Procedure_Results_Ranked"],
                ["18", "IRI Ranked Summary", "IRI_Ranked_Summary"],
                ["19", "IRI Triggered Detail", "IRI_Detail"],
            ]);

        mainPart.Document.Save();
    }

    private static void WriteTestSection(Body body, string heading, string objective, string excelSheet, List<MasterRow> rows, bool applicable)
    {
        body.Append(Heading(heading, 2));
        body.Append(Paragraph($"Objective: {objective}"));

        if (!applicable)
        {
            body.Append(Paragraph("Status: Not Applicable for this run configuration."));
            body.Append(Paragraph($"Detailed output in Excel sheet '{excelSheet}' is not generated when the procedure is set to Not Applicable."));
            AddAuditorCommentBlock(body, "Reason for Not Applicable decision and impact on audit scope.");
            return;
        }

        var journalCount = DistinctJournalCount(rows, useGroupKey: string.Equals(excelSheet, "Unbalanced", StringComparison.OrdinalIgnoreCase));
        var totalAbs = rows.Sum(x => x.AbsAmount);
        body.Append(Paragraph($"Result: {journalCount:N0} journal groups ({rows.Count:N0} lines)."));
        AddSimpleTable(body,
            ["Metric", "Result"],
            [
                ["Journal Groups", journalCount.ToString("N0")],
                ["Lines", rows.Count.ToString("N0")],
                ["Total Absolute Value", $"R{totalAbs:N2}"],
            ]);

        if (rows.Count == 0)
        {
            body.Append(Paragraph("No exceptions identified."));
            body.Append(Paragraph($"Detailed population is available in Excel sheet: {excelSheet}."));
            AddAuditorCommentBlock(body, "Conclusion and corroborating evidence for no exceptions.");
            return;
        }

        AddSimpleTable(body,
            ["Top Account", "Account Name/Description", "Lines", "Journal Groups", "Total Abs Value"],
            rows
                .GroupBy(r => r.Acct)
                .Select(g => new
                {
                    Account = g.Key,
                    Description = g.Select(x => x.Description).FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? string.Empty,
                    Lines = g.Count(),
                    Journals = g.Select(x => x.JournalCompositeKey).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
                    TotalAbs = g.Sum(x => x.AbsAmount),
                })
                .OrderByDescending(x => x.TotalAbs)
                .Take(10)
                .Select(x => new[]
                {
                    x.Account,
                    x.Description,
                    x.Lines.ToString("N0"),
                    x.Journals.ToString("N0"),
                    $"R{x.TotalAbs:N2}",
                }),
            10);
        body.Append(Paragraph($"Detailed line-level results are available in Excel sheet: {excelSheet}."));
        AddAuditorCommentBlock(body, "Interpretation of results, root causes, and agreed management actions.");
    }

    private static string[] BuildOverviewRow(string procedure, bool applicable, int journals, int lines, string excelSheet)
    {
        if (!applicable)
        {
            return [procedure, "Not Applicable", "N/A", "N/A", excelSheet];
        }

        return [procedure, "Applicable", journals.ToString("N0"), lines.ToString("N0"), excelSheet];
    }

    private static int DistinctJournalCount(List<MasterRow> rows, bool useGroupKey = false)
    {
        var keySelector = useGroupKey
            ? rows.Select(x => x.JournalKey)
            : rows.Select(x => x.JournalCompositeKey);

        return keySelector
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Count();
    }

    private static IEnumerable<string[]> BuildDayBreakdownRows(IEnumerable<MasterRow> rows)
    {
        return rows
            .GroupBy(r => string.IsNullOrWhiteSpace(r.DayName) ? "Unknown" : r.DayName)
            .Select(g => new
            {
                Day = g.Key,
                Lines = g.Count(),
                Journals = g.Select(x => x.JournalCompositeKey).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
            })
            .OrderByDescending(x => x.Lines)
            .ThenBy(x => x.Day, StringComparer.OrdinalIgnoreCase)
            .Select(x => new[]
            {
                x.Day,
                x.Lines.ToString("N0"),
                x.Journals.ToString("N0"),
            });
    }

    private static IEnumerable<string[]> BuildHolidayBreakdownRows(IEnumerable<MasterRow> rows)
    {
        return rows
            .Where(r => !string.IsNullOrWhiteSpace(r.HolidayName))
            .GroupBy(r => new
            {
                Holiday = r.HolidayName,
                Day = string.IsNullOrWhiteSpace(r.DayName) ? "Unknown" : r.DayName,
            })
            .Select(g => new
            {
                g.Key.Holiday,
                g.Key.Day,
                Lines = g.Count(),
                Journals = g.Select(x => x.JournalCompositeKey).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
            })
            .OrderByDescending(x => x.Lines)
            .ThenBy(x => x.Holiday, StringComparer.OrdinalIgnoreCase)
            .Select(x => new[]
            {
                x.Holiday,
                x.Day,
                x.Lines.ToString("N0"),
                x.Journals.ToString("N0"),
            });
    }

    private static void AddAuditorCommentBlock(Body body, string guidance)
    {
        body.Append(Paragraph("Auditor Comment:"));
        body.Append(Paragraph($"Guidance: {guidance}"));
        body.Append(Paragraph(".............................................................................................................................."));
        body.Append(Paragraph(".............................................................................................................................."));
        body.Append(Paragraph(".............................................................................................................................."));
    }

    private static void AddSimpleTable(Body body, string[] headers, IEnumerable<string[]> rows, int maxRows = 1000)
    {
        var table = new Table();
        table.AppendChild(new TableProperties(
            new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" },
            new TableLayout { Type = TableLayoutValues.Fixed },
            new TableJustification { Val = TableRowAlignmentValues.Center },
            new TableBorders(
                new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 })));

        var headerRow = new TableRow();
        foreach (var h in headers)
        {
            var cell = new TableCell(new Paragraph(new Run(new Text(h ?? string.Empty))));
            cell.Append(new TableCellProperties(
                new Shading
                {
                    Val = ShadingPatternValues.Clear,
                    Color = "auto",
                    Fill = "CDB7E9",
                }));

            foreach (var p in cell.Elements<Paragraph>())
            {
                foreach (var r in p.Elements<Run>())
                {
                    r.PrependChild(new RunProperties(
                        new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
                        new Bold(),
                        new Color { Val = "1B2A6B" },
                        new FontSize { Val = "19" }));
                }
            }

            headerRow.Append(cell);
        }

        table.Append(headerRow);

        var idx = 0;
        foreach (var row in rows.Take(maxRows))
        {
            var tr = new TableRow();
            var fill = idx % 2 == 0 ? "FFFFFF" : "F6F2FC";
            foreach (var col in row)
            {
                var cell = new TableCell(new Paragraph(new Run(new Text(col ?? string.Empty))));
                cell.Append(new TableCellProperties(
                    new Shading
                    {
                        Val = ShadingPatternValues.Clear,
                        Color = "auto",
                        Fill = fill,
                    }));

                foreach (var p in cell.Elements<Paragraph>())
                {
                    foreach (var r in p.Elements<Run>())
                    {
                        r.PrependChild(new RunProperties(
                            new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
                            new FontSize { Val = "18" },
                            new Color { Val = "222222" }));
                    }
                }

                tr.Append(cell);
            }

            table.Append(tr);
            idx++;
        }

        body.Append(table);
    }

    private static Paragraph Heading(string text, int level)
    {
        var p = new Paragraph();
        var styleId = level == 1 ? "Heading1" : "Heading2";
        var props = new ParagraphProperties(
            new ParagraphStyleId { Val = styleId },
            new OutlineLevel { Val = level == 1 ? 0 : 1 },
            new SpacingBetweenLines
            {
                Before = level == 1 ? "240" : "160",
                After = level == 1 ? "120" : "80",
            });
        p.Append(props);

        var run = new Run(new Text(text));
        run.PrependChild(new RunProperties(
            new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
            new Bold(),
            new Color { Val = level == 1 ? "1B2A6B" : "4F2D7F" },
            new FontSize { Val = level == 1 ? "34" : "26" }));
        p.Append(run);

        return p;
    }

    private static void AddTableOfContents(Body body)
    {
        body.Append(Heading("Contents", 2));

        var toc = new Paragraph(
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode("TOC \\o \"1-3\" \\h \\z \\u") { Space = SpaceProcessingModeValues.Preserve }),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new Run(new Text("Right-click and choose 'Update Field' in Word to refresh the table of contents.")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        toc.PrependChild(new ParagraphProperties(
            new SpacingBetweenLines
            {
                Before = "40",
                After = "160",
            }));
        body.Append(toc);

        // Start report body after TOC so Word can paginate and align sections clearly.
        body.Append(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
    }

    private static void AddJournalCaatObjectiveSection(Body body, CaatsState state)
    {
        body.Append(Heading("2. Journal CAAT Objective", 2));

        var glSystem = string.IsNullOrWhiteSpace(state.Engagement.GlSystem) ? "client GL" : state.Engagement.GlSystem;
        AddSimpleTable(body,
            ["CAAT", "Description"],
            [
                [
                    "Journal Completeness",
                    "Assess completeness of journal entries in the financial statements by obtaining trial balance movements and comparing them to journal movement by GL account."
                ],
                [
                    "Weekend and Public Holiday Manual Journals",
                    "Identify manual journals entered on weekends and/or public holidays."
                ],
                [
                    "User Analysis",
                    "Analyze users posting journals throughout the financial period, including manual and automated behavior."
                ],
                [
                    "Unbalanced Journals",
                    "Identify journals where debit and credit totals do not balance at journal-document level."
                ],
                [
                    "Journals above performance materiality",
                    "Identify journals above performance materiality and summarize population impact."
                ],
                [
                    "Account Analysis",
                    "Identify most and least used accounts based on number of journal lines processed per GL account."
                ],
                [
                    "Round Amount Analysis",
                    "Identify rounded patterns such as trailing 0 and trailing 9 amounts (with and without decimals)."
                ],
                [
                    "Backdated Journal Entries",
                    "Identify journals where creation date period is later than posting date period and quantify day differences."
                ],
            ]);

        body.Append(Paragraph($"Procedures are executed against {glSystem} data using configured CAAT rules and engagement thresholds."));
    }

    private static void AddAnalysisStepsSection(Body body, CaatsState state)
    {
        body.Append(Heading("3. Analysis Steps and Audit Procedures", 2));

        var origins = string.IsNullOrWhiteSpace(state.Engagement.SngOriginsRaw)
            ? "FAJ, GJ, IJ, Journal, OBJ, PYRJ"
            : state.Engagement.SngOriginsRaw;

        AddSimpleTable(body,
            ["ID", "Description"],
            [
                ["1", "Obtain GL journal extracts from the client/audit team."],
                ["2", "Obtain trial balance from the audit team."],
                ["3", $"Apply SNG Grant Thornton manual-journal rule in {state.Engagement.GlSystem}: journal origin in ({origins})."],
                ["4", "Validate mapping, period filters, and completeness of required fields."],
                ["5", "Process journals, identify exceptions, and categorize using SNG Grant Thornton methodology."],
                ["6", "Perform scoped procedure analytics (weekend, holiday, backdated, adj/correc, materiality, round, duplicate, unbalanced, low FSLI)."],
                ["7", "Generate summary working paper and detailed Excel annexures for audit evidence and review."],
            ]);
    }

    private static Paragraph Paragraph(string text)
    {
        var run = new Run(new Text(text ?? string.Empty));
        run.PrependChild(new RunProperties(
            new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
            new FontSize { Val = "21" },
            new Color { Val = "2A2A2A" }));

        var p = new Paragraph(run);
        p.PrependChild(new ParagraphProperties(
            new SpacingBetweenLines
            {
                Before = "20",
                After = "60",
            }));
        return p;
    }

    private static void AppendTitleBlock(Body body, CaatsState state, string reportDate)
    {
        var title = new Paragraph();
        title.PrependChild(new ParagraphProperties(
            new Justification { Val = JustificationValues.Center },
            new SpacingBetweenLines { Before = "160", After = "90" }));
        var titleRun = new Run(new Text("JOURNAL WORKING PAPER"));
        titleRun.PrependChild(new RunProperties(
            new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
            new Bold(),
            new Color { Val = "1B2A6B" },
            new FontSize { Val = "44" }));
        title.Append(titleRun);
        body.Append(title);

        var subtitle = new Paragraph();
        subtitle.PrependChild(new ParagraphProperties(
            new Justification { Val = JustificationValues.Center },
            new SpacingBetweenLines { Before = "30", After = "150" }));
        var subtitleRun = new Run(new Text("Computer Assisted Audit Techniques (CAATs) - SNG Grant Thornton"));
        subtitleRun.PrependChild(new RunProperties(
            new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
            new Color { Val = "4F2D7F" },
            new FontSize { Val = "24" }));
        subtitle.Append(subtitleRun);
        body.Append(subtitle);

        AddSimpleTable(body,
            ["Engagement", "Client", "GL System", "Report Date"],
            [[state.Engagement.Engagement, state.Engagement.Client, state.Engagement.GlSystem, reportDate]],
            1);
    }

    private static IEnumerable<string> BuildMasterCsv(IEnumerable<MasterRow> rows)
    {
        yield return "Account,JournalKey,JournalId,PostingDate,CpuDate,Debit,Credit,Signed,AbsAmount,ManualReason,TestBackdated,TestRound,TestDuplicate,TestUnbalanced";
        foreach (var r in rows)
        {
            yield return string.Join(',', [
                Csv(r.Acct), Csv(r.JournalKey), Csv(r.JournalId), Csv(r.PostingDate?.ToString("yyyy-MM-dd") ?? string.Empty),
                Csv(r.CpuDate?.ToString("yyyy-MM-dd") ?? string.Empty), r.Debit.ToString("0.00"), r.Credit.ToString("0.00"),
                r.Signed.ToString("0.00"), r.AbsAmount.ToString("0.00"), Csv(r.ManualReason), r.TestBackdated.ToString(),
                r.TestRound.ToString(), r.TestDuplicate.ToString(), r.TestUnbalanced.ToString(),
            ]);
        }
    }

    private static IEnumerable<string> BuildReconCsv(IEnumerable<ReconRow> rows)
    {
        yield return "AccountNumber,AccountName,OpeningBalance,ClosingBalance,GLDebit,GLCredit,GLBalance,Difference,Status,GLSource";
        foreach (var r in rows)
        {
            yield return string.Join(',', [
                Csv(r.AccountNumber), Csv(r.AccountName), r.OpeningBalance.ToString("0.00"), r.ClosingBalance.ToString("0.00"),
                r.GlDebit.ToString("0.00"), r.GlCredit.ToString("0.00"), r.GlBalance.ToString("0.00"), r.Difference.ToString("0.00"), Csv(r.Status), Csv(r.GlSource ?? string.Empty),
            ]);
        }
    }

    private static string Csv(string s)
    {
        if (s.Contains(',') || s.Contains('"') || s.Contains('\n'))
        {
            return $"\"{s.Replace("\"", "\"\"", StringComparison.Ordinal)}\"";
        }

        return s;
    }
}
