namespace CaatsWebApp.Models.Caats;

public sealed class TablePreviewResponse
{
    public List<string> GlColumns { get; set; } = [];
    public List<string> TbColumns { get; set; } = [];
    public List<Dictionary<string, object?>> GlSample { get; set; } = [];
    public List<Dictionary<string, object?>> TbSample { get; set; } = [];
}

public sealed class GlProfileResponse
{
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
    public int TotalRows { get; set; }
    public string DateColumnUsed { get; set; } = string.Empty;
    public DateTime? MinDate { get; set; }
    public DateTime? MaxDate { get; set; }
    public int? FyYear { get; set; }
    public decimal WeekendPercent { get; set; }
    public bool WeekendNormalSuggested { get; set; }
    public string AmountStructure { get; set; } = "unknown";
    public string SuggestedIndustry { get; set; } = "Standard (office hours)";
}

public sealed class RunTestsResponse
{
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
    public Dictionary<string, int> Counts { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public int ReconAgree { get; set; }
    public int ReconVariance { get; set; }
    public decimal TotalDebit { get; set; }
    public decimal TotalCredit { get; set; }
    public int BenfordFlaggedDigits { get; set; }
    public int ReconTotalRows { get; set; }
    public List<BenfordRow> Benford { get; set; } = [];
    public List<ReconRow> ReconTop { get; set; } = [];
    public List<ReconRow> ReconRows { get; set; } = [];
    public List<string> Users { get; set; } = [];
    public List<string> Accounts { get; set; } = [];
    public List<string> JournalKeys { get; set; } = [];
    public decimal TotalRiskScore { get; set; }
    public List<RiskScoreSummaryRow> RiskScoreSummary { get; set; } = [];
    public List<UserAnalysisFullRow> UserAnalysisFull { get; set; } = [];
    public List<ProcedureRankedRow> ProcedureResultsRanked { get; set; } = [];
    public List<IriRankedRow> IriRankedSummary { get; set; } = [];
    public List<IriDetailRow> IriDetailRows { get; set; } = [];
    public bool HolidayProcedureApplicable { get; set; }
    public List<HolidayVerificationRow> HolidayVerification { get; set; } = [];
}

public sealed class HolidayVerificationRow
{
    public DateTime CheckDate { get; set; }
    public string HolidayName { get; set; } = string.Empty;
    public int MatchedLines { get; set; }
    public int MatchedJournalGroups { get; set; }
    public int InScopeLines { get; set; }
    public int InScopeJournalGroups { get; set; }
    public int InScopeExceptionLines { get; set; }
    public int InScopeExceptionJournalGroups { get; set; }
}

public sealed class RiskScoreSummaryRow
{
    public string Test { get; set; } = string.Empty;
    public decimal TotalRiskScore { get; set; }
    public int LinesFlagged { get; set; }
}

public sealed class UserAnalysisFullRow
{
    public string FullName { get; set; } = string.Empty;
    public string ManualAutomatedDescriptor { get; set; } = string.Empty;
    public int LineCountDebit { get; set; }
    public decimal TotalReportingAmountDebit { get; set; }
    public int LineCountCredit { get; set; }
    public decimal TotalReportingAmountCredit { get; set; }
    public int TotalLineCount { get; set; }
    public decimal TotalReportingAmount { get; set; }
}

public sealed class ProcedureRankedRow
{
    public int Rank { get; set; }
    public string Procedure { get; set; } = string.Empty;
    public int Count { get; set; }
    public string RankedLabel { get; set; } = string.Empty;
}

public sealed class IriRankedRow
{
    public int Rank { get; set; }
    public string Iri { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public int Count { get; set; }
    public string RankedLabel { get; set; } = string.Empty;
}

public sealed class IriDetailRow
{
    public string Iri { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string Acct { get; set; } = string.Empty;
    public string JournalKey { get; set; } = string.Empty;
    public string JournalId { get; set; } = string.Empty;
    public DateTime? PostingDate { get; set; }
    public DateTime? CpuDate { get; set; }
    public string User { get; set; } = string.Empty;
    public decimal Debit { get; set; }
    public decimal Credit { get; set; }
    public decimal AbsAmount { get; set; }
    public string DayName { get; set; } = string.Empty;
    public string HolidayName { get; set; } = string.Empty;
    public int MonthsBackdated { get; set; }
    public string RoundCategory { get; set; } = string.Empty;
}

public sealed class ExploreRequest
{
    public string Filter { get; set; } = "All Records";
    public string User { get; set; } = "All";
    public string Account { get; set; } = "All";
    public string JournalKey { get; set; } = "All";
    public int MaxRows { get; set; } = 100;
}

public sealed class ExploreResponse
{
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
    public int Journals { get; set; }
    public int Lines { get; set; }
    public List<MasterRow> Rows { get; set; } = [];
}

public sealed class MasterRow
{
    public string Acct { get; set; } = string.Empty;
    public string JournalKey { get; set; } = string.Empty;
    public string JournalCompositeKey { get; set; } = string.Empty;
    public string JournalId { get; set; } = string.Empty;
    public DateTime? PostingDate { get; set; }
    public DateTime? CpuDate { get; set; }
    public DateTime? CheckDate { get; set; }
    public string DayName { get; set; } = string.Empty;
    public string HolidayName { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string User { get; set; } = "Unknown";
    public decimal Debit { get; set; }
    public decimal Credit { get; set; }
    public decimal Signed { get; set; }
    public decimal AbsAmount { get; set; }
    public string JnlOrigin { get; set; } = string.Empty;
    public int IsManual { get; set; }
    public string ManualReason { get; set; } = string.Empty;
    public int MonthsBackdated { get; set; }
    public int DaysBackdated { get; set; }
    public string RoundCategory { get; set; } = "Other";
    public string RoundPattern { get; set; } = "Other";
    public int? BenfordDigit { get; set; }
    public string Period { get; set; } = string.Empty;
    public int TestWeekend { get; set; }
    public int TestHoliday { get; set; }
    public int TestWeekendInfo { get; set; }
    public int TestHolidayInfo { get; set; }
    public int TestWeekendManual { get; set; }
    public int TestHolidayManual { get; set; }
    public int TestBackdated { get; set; }
    public int TestAdjDesc { get; set; }
    public int TestAboveMat { get; set; }
    public int TestRound { get; set; }
    public int TestDuplicate { get; set; }
    public int TestUnbalanced { get; set; }
    public int TestLowFsli { get; set; }
}

public sealed class ReconRow
{
    public string AccountNumber { get; set; } = string.Empty;
    public string AccountName { get; set; } = string.Empty;
    public decimal OpeningBalance { get; set; }
    public decimal ClosingBalance { get; set; }
    public decimal GlDebit { get; set; }
    public decimal GlCredit { get; set; }
    public decimal GlBalance { get; set; }
    public decimal Difference { get; set; }
    public string Status { get; set; } = string.Empty;
    public string GlSource { get; set; } = string.Empty;
}

public sealed class BenfordRow
{
    public int LeadingDigit { get; set; }
    public int Count { get; set; }
    public decimal ActualPercent { get; set; }
    public decimal ExpectedPercent { get; set; }
    public decimal DifferencePercent { get; set; }
    public string Status { get; set; } = "";
}
