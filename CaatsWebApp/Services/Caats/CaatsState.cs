using System.Data;
using CaatsWebApp.Models.Caats;

namespace CaatsWebApp.Services.Caats;

public sealed class CaatsState
{
    public string Server { get; set; } = string.Empty;
    public string Database { get; set; } = string.Empty;
    public string ConnectionString { get; set; } = string.Empty;
    public string GlTable { get; set; } = string.Empty;
    public string TbTable { get; set; } = string.Empty;
    public DataTable? GlData { get; set; }
    public DataTable? TbData { get; set; }
    public string LoadedDatabase { get; set; } = string.Empty;
    public string LoadedGlTable { get; set; } = string.Empty;
    public string LoadedTbTable { get; set; } = string.Empty;
    public Dictionary<string, string> GlMap { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, string> TbMap { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public EngagementSettings Engagement { get; set; } = new();
    public List<MasterRow> MasterRows { get; set; } = [];
    public List<ReconRow> ReconRows { get; set; } = [];
    public List<BenfordRow> BenfordRows { get; set; } = [];
    public List<RiskScoreSummaryRow> RiskScoreSummaryRows { get; set; } = [];
    public List<UserAnalysisFullRow> UserAnalysisFullRows { get; set; } = [];
    public List<ProcedureRankedRow> ProcedureRankedRows { get; set; } = [];
    public List<IriRankedRow> IriRankedRows { get; set; } = [];
    public List<IriDetailRow> IriDetailRows { get; set; } = [];
    public decimal TotalRiskScore { get; set; }
    public Dictionary<string, bool> LastProcedures { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, ProcedureScope> LastProcedureScopes { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}
