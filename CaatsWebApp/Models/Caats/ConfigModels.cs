namespace CaatsWebApp.Models.Caats;

public sealed class TableSelectionRequest
{
    public string GlTable { get; set; } = string.Empty;
    public string TbTable { get; set; } = string.Empty;
}

public sealed class ColumnMapRequest
{
    public Dictionary<string, string> Gl { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, string> Tb { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

public sealed class EngagementSettings
{
    public string Parent { get; set; } = string.Empty;
    public string Client { get; set; } = "Client";
    public string Engagement { get; set; } = string.Empty;
    public string GlSystem { get; set; } = "Custom / Other";
    public string Industry { get; set; } = "Standard (office hours)";
    public bool WeekendNormal { get; set; }
    public bool HolidayNormal { get; set; }
    public bool SngRule { get; set; }
    public string SngOriginsRaw { get; set; } = "FAJ,GJ,IJ,Journal,OBJ,PYRJ";
    public string JournalGroupColumn { get; set; } = string.Empty;
    public string WeekendDateColumn { get; set; } = string.Empty;
    public string GlReconAmountColumn { get; set; } = string.Empty;
    public DateTime? FyStart { get; set; }
    public DateTime? FyEnd { get; set; }
    public DateTime? PeriodStart { get; set; }
    public DateTime? PeriodEnd { get; set; }
    public DateTime? SigningDate { get; set; }
    public decimal Materiality { get; set; } = 10_000_000m;
    public decimal PerformanceMateriality { get; set; } = 7_500_000m;
    public decimal MinJournalValue { get; set; }
    public int LowFsliThreshold { get; set; } = 10;
    public string Auditor { get; set; } = string.Empty;
    public string Manager { get; set; } = string.Empty;
    public string CountryCode { get; set; } = "ZA";
}

public sealed class RunTestsRequest
{
    public Dictionary<string, bool> Procedures { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, ProcedureScope> ProcedureScopes { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

public sealed class ProcedureScope
{
    public bool Manual { get; set; } = true;
    public bool SemiAuto { get; set; } = true;
    public bool Auto { get; set; } = true;
}
