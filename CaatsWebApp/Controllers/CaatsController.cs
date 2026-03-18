using CaatsWebApp.Models.Caats;
using CaatsWebApp.Services.Caats;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.StaticFiles;
using System.Globalization;
using System.Data;
using System.IO;

namespace CaatsWebApp.Controllers;

[ApiController]
[Route("api/[controller]")]
public sealed class CaatsController : ControllerBase
{
    private readonly CaatsState _state;
    private readonly SqlDataService _sql;
    private readonly CaatsAnalysisService _analysis;
    private readonly ExportService _export;

    public CaatsController(CaatsState state, SqlDataService sql, CaatsAnalysisService analysis, ExportService export)
    {
        _state = state;
        _sql = sql;
        _analysis = analysis;
        _export = export;
    }

    [HttpGet("drivers")]
    public IActionResult Drivers()
    {
        return Ok(new[]
        {
            "ODBC Driver 18 for SQL Server",
            "ODBC Driver 17 for SQL Server",
            "SQL Server",
        });
    }

    [HttpGet("default-output-folder")]
    public IActionResult DefaultOutputFolder()
    {
        var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        if (string.IsNullOrWhiteSpace(desktop))
        {
            desktop = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Desktop");
        }

        return Ok(new
        {
            outputFolder = Path.Combine(desktop, "JE_Audit"),
        });
    }

    [HttpPost("connect")]
    public async Task<ActionResult<ConnectResponse>> Connect([FromBody] ConnectRequest req)
    {
        try
        {
            var auth = req.TrustedConnection
                ? "Integrated Security=True;TrustServerCertificate=True;"
                : $"User ID={req.Username};Password={req.Password};TrustServerCertificate=True;";

            var cs = $"Server={req.Server};Database=master;{auth}";
            var dbs = await _sql.GetDatabasesAsync(cs);
            _state.Server = req.Server;
            _state.ConnectionString = cs;

            return Ok(new ConnectResponse
            {
                Success = true,
                Message = $"Connected. {dbs.Count} databases found.",
                Databases = dbs,
            });
        }
        catch (Exception ex)
        {
            return BadRequest(new ConnectResponse { Success = false, Message = ex.Message });
        }
    }

    [HttpPost("database")]
    public async Task<ActionResult<GenericResponse>> UseDatabase([FromBody] UseDatabaseRequest req)
    {
        if (string.IsNullOrWhiteSpace(_state.Server))
        {
            return BadRequest(new GenericResponse { Success = false, Message = "Connect first." });
        }

        try
        {
            var builder = new Microsoft.Data.SqlClient.SqlConnectionStringBuilder(_state.ConnectionString)
            {
                InitialCatalog = req.Database,
            };
            _state.Database = req.Database;
            _state.ConnectionString = builder.ConnectionString;
            _state.GlData = null;
            _state.TbData = null;
            _state.LoadedDatabase = string.Empty;
            _state.LoadedGlTable = string.Empty;
            _state.LoadedTbTable = string.Empty;
            var tables = await _sql.GetTablesAsync(_state.ConnectionString);
            return Ok(new GenericResponse { Success = true, Message = $"Database selected. {tables.Count} tables available." });
        }
        catch (Exception ex)
        {
            return BadRequest(new GenericResponse { Success = false, Message = ex.Message });
        }
    }

    [HttpGet("tables")]
    public async Task<IActionResult> Tables()
    {
        if (string.IsNullOrWhiteSpace(_state.ConnectionString) || string.IsNullOrWhiteSpace(_state.Database))
        {
            return BadRequest(new GenericResponse { Success = false, Message = "Select a database first." });
        }

        try
        {
            var tables = await _sql.GetTablesAsync(_state.ConnectionString);
            return Ok(tables);
        }
        catch (Exception ex)
        {
            return BadRequest(new GenericResponse { Success = false, Message = ex.Message });
        }
    }

    [HttpPost("preview")]
    public async Task<ActionResult<TablePreviewResponse>> Preview([FromBody] TableSelectionRequest req)
    {
        if (string.IsNullOrWhiteSpace(_state.ConnectionString))
        {
            return BadRequest(new GenericResponse { Success = false, Message = "Connect first." });
        }

        try
        {
            _state.GlTable = req.GlTable;
            _state.TbTable = req.TbTable;

            // Table selection changed; force explicit reload for full datasets.
            _state.GlData = null;
            _state.TbData = null;
            _state.LoadedDatabase = string.Empty;
            _state.LoadedGlTable = string.Empty;
            _state.LoadedTbTable = string.Empty;

            var glTask = _sql.GetTableAsync(_state.ConnectionString, req.GlTable, 5);
            var tbTask = _sql.GetTableAsync(_state.ConnectionString, req.TbTable, 5);
            await Task.WhenAll(glTask, tbTask);

            var gl = glTask.Result;
            var tb = tbTask.Result;

            return Ok(new TablePreviewResponse
            {
                GlColumns = gl.Columns.Cast<System.Data.DataColumn>().Select(c => c.ColumnName).ToList(),
                TbColumns = tb.Columns.Cast<System.Data.DataColumn>().Select(c => c.ColumnName).ToList(),
                GlSample = gl.Rows.Cast<DataRow>().Select(r => gl.Columns.Cast<System.Data.DataColumn>().ToDictionary(c => c.ColumnName, c => r[c] == DBNull.Value ? null : r[c])).ToList(),
                TbSample = tb.Rows.Cast<DataRow>().Select(r => tb.Columns.Cast<System.Data.DataColumn>().ToDictionary(c => c.ColumnName, c => r[c] == DBNull.Value ? null : r[c])).ToList(),
            });
        }
        catch (Exception ex)
        {
            return BadRequest(new GenericResponse { Success = false, Message = ex.Message });
        }
    }

    [HttpPost("mapping")]
    public IActionResult SaveMapping([FromBody] ColumnMapRequest req)
    {
        _state.GlMap = new Dictionary<string, string>(req.Gl, StringComparer.OrdinalIgnoreCase);
        _state.TbMap = new Dictionary<string, string>(req.Tb, StringComparer.OrdinalIgnoreCase);
        return Ok(new GenericResponse { Success = true, Message = "Column mapping saved." });
    }

    [HttpPost("load-data")]
    public async Task<IActionResult> LoadData()
    {
        if (string.IsNullOrWhiteSpace(_state.GlTable) || string.IsNullOrWhiteSpace(_state.TbTable))
        {
            return BadRequest(new GenericResponse { Success = false, Message = "Select GL/TB tables first." });
        }

        try
        {
            // Reuse in-memory datasets when the selected tables are already loaded.
            if (_state.GlData is not null && _state.TbData is not null &&
                _state.GlData.Rows.Count > 0 && _state.TbData.Rows.Count > 0 &&
                string.Equals(_state.LoadedDatabase, _state.Database, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(_state.LoadedGlTable, _state.GlTable, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(_state.LoadedTbTable, _state.TbTable, StringComparison.OrdinalIgnoreCase))
            {
                return Ok(new
                {
                    success = true,
                    glRows = _state.GlData.Rows.Count,
                    tbRows = _state.TbData.Rows.Count,
                    message = "Data already loaded from memory.",
                });
            }

            var glTask = _sql.GetTableAsync(_state.ConnectionString, _state.GlTable);
            var tbTask = _sql.GetTableAsync(_state.ConnectionString, _state.TbTable);
            await Task.WhenAll(glTask, tbTask);

            _state.GlData = glTask.Result;
            _state.TbData = tbTask.Result;
            _state.LoadedDatabase = _state.Database;
            _state.LoadedGlTable = _state.GlTable;
            _state.LoadedTbTable = _state.TbTable;
            return Ok(new
            {
                success = true,
                glRows = _state.GlData.Rows.Count,
                tbRows = _state.TbData.Rows.Count,
                message = "Data loaded.",
            });
        }
        catch (Exception ex)
        {
            return BadRequest(new GenericResponse { Success = false, Message = ex.Message });
        }
    }

    [HttpPost("profile")]
    public ActionResult<GlProfileResponse> Profile()
    {
        if (_state.GlData is null || _state.GlData.Rows.Count == 0)
        {
            return BadRequest(new GlProfileResponse
            {
                Success = false,
                Message = "Load full GL data first (Step 2: Load Full Data).",
            });
        }

        var dateColumn = ResolveDateColumn(_state.GlData, _state.Engagement, _state.GlMap);
        var dates = new List<DateTime>();
        if (!string.IsNullOrWhiteSpace(dateColumn))
        {
            foreach (DataRow row in _state.GlData.Rows)
            {
                if (row[dateColumn!] == DBNull.Value)
                {
                    continue;
                }

                if (TryDate(row[dateColumn!], out var parsed))
                {
                    dates.Add(parsed.Date);
                }
            }
        }

        var weekendPct = 0m;
        if (dates.Count > 0)
        {
            var weekendCount = dates.Count(d => d.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday);
            weekendPct = decimal.Round((decimal)weekendCount / dates.Count * 100m, 1);
        }

        var weekendNormalSuggested = weekendPct > 15m;
        var industry = weekendNormalSuggested ? "Retail (24/7 operations)" : "Standard (office hours)";

        var hasDebit = HasColumnLike(_state.GlData, _state.GlMap.GetValueOrDefault("debit"), "debit", "dr");
        var hasCredit = HasColumnLike(_state.GlData, _state.GlMap.GetValueOrDefault("credit"), "credit", "cr");
        var hasSigned = HasColumnLike(_state.GlData, _state.GlMap.GetValueOrDefault("amount_signed"), "general ledger amt", "amt in loc.cur", "dmbtr", "signed", "amount");
        var hasBalance = HasColumnLike(_state.GlData, _state.Engagement.GlReconAmountColumn, "balance", "amount");

        var amountStructure = hasDebit && hasCredit
            ? "debit_credit"
            : hasSigned
                ? "signed_col"
            : hasBalance
                ? "balance_col"
                : "unknown";

        return Ok(new GlProfileResponse
        {
            Success = true,
            Message = "GL profile detected.",
            TotalRows = _state.GlData.Rows.Count,
            DateColumnUsed = dateColumn ?? string.Empty,
            MinDate = dates.Count > 0 ? dates.Min() : null,
            MaxDate = dates.Count > 0 ? dates.Max() : null,
            FyYear = dates.Count > 0 ? dates.Max().Year : null,
            WeekendPercent = weekendPct,
            WeekendNormalSuggested = weekendNormalSuggested,
            AmountStructure = amountStructure,
            SuggestedIndustry = industry,
        });
    }

    [HttpPost("engagement")]
    public IActionResult SaveEngagement([FromBody] EngagementSettings req)
    {
        if (string.Equals(req.GlSystem, "SAP FI", StringComparison.OrdinalIgnoreCase)
            || string.Equals(req.GlSystem, "SAP", StringComparison.OrdinalIgnoreCase))
        {
            req.SngRule = true;
            req.SngOriginsRaw = "SA";
            if (string.IsNullOrWhiteSpace(req.GlReconAmountColumn))
            {
                // SAP FI recon should default to signed amount SUMIF behavior.
                req.GlReconAmountColumn = string.Empty;
            }
        }

        _state.Engagement = req;
        return Ok(new GenericResponse { Success = true, Message = "Engagement saved." });
    }

    [HttpPost("run")]
    public ActionResult<RunTestsResponse> Run([FromBody] RunTestsRequest req)
    {
        if (_state.GlData is null || _state.TbData is null)
        {
            return BadRequest(new RunTestsResponse { Success = false, Message = "Load data first." });
        }

        try
        {
            var (master, recon, benford) = _analysis.Run(_state.GlData, _state.TbData, _state.GlMap, _state.TbMap, _state.Engagement, _state.GlTable, _state.Database);
            _state.MasterRows = master;
            _state.ReconRows = recon;
            _state.BenfordRows = benford;

            var riskSummary = BuildRiskScoreSummary(master);
            var totalRiskScore = riskSummary.Where(x => !string.Equals(x.Test, "TOTAL", StringComparison.OrdinalIgnoreCase)).Sum(x => x.TotalRiskScore);
            riskSummary.Add(new RiskScoreSummaryRow
            {
                Test = "TOTAL",
                TotalRiskScore = decimal.Round(totalRiskScore, 2),
                LinesFlagged = master.Count,
            });

            var userAnalysisFull = BuildUserAnalysisFull(master);
            var procedureRanked = BuildProcedureRanked(master, recon);
            var iriRanked = BuildIriRanked(master, _state.Engagement);
            var iriDetails = BuildIriDetails(master, _state.Engagement);

            _state.RiskScoreSummaryRows = riskSummary;
            _state.TotalRiskScore = decimal.Round(totalRiskScore, 2);
            _state.UserAnalysisFullRows = userAnalysisFull;
            _state.ProcedureRankedRows = procedureRanked;
            _state.IriRankedRows = iriRanked;
            _state.IriDetailRows = iriDetails;

            var procedures = req.Procedures ?? new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            var procedureScopes = req.ProcedureScopes ?? new Dictionary<string, ProcedureScope>(StringComparer.OrdinalIgnoreCase);
            _state.LastProcedures = new Dictionary<string, bool>(procedures, StringComparer.OrdinalIgnoreCase);
            _state.LastProcedureScopes = new Dictionary<string, ProcedureScope>(procedureScopes, StringComparer.OrdinalIgnoreCase);
            static bool IsApplicable(Dictionary<string, bool> map, string key)
            {
                return !map.TryGetValue(key, out var enabled) || enabled;
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

            int CountScopedJournalGroups(bool applicable, string procedureKey, Func<MasterRow, bool> flagged)
            {
                if (!applicable)
                {
                    return -1;
                }

                var scope = procedureScopes.TryGetValue(procedureKey, out var s) ? s : null;
                var scopedFlagged = master.Where(r => InScope(r, scope) && flagged(r));

                if (string.Equals(procedureKey, "unbalanced", StringComparison.OrdinalIgnoreCase))
                {
                    return scopedFlagged
                        .Select(r => r.JournalKey)
                        .Where(k => !string.IsNullOrWhiteSpace(k))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .Count();
                }

                return scopedFlagged
                    .Select(r => r.JournalCompositeKey)
                    .Where(k => !string.IsNullOrWhiteSpace(k))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Count();
            }

            var weekendApplicable = IsApplicable(procedures, "weekend");
            var holidayApplicable = IsApplicable(procedures, "holiday");
            var backdatedApplicable = IsApplicable(procedures, "backdated");
            var adjApplicable = IsApplicable(procedures, "adj_desc");
            var matApplicable = IsApplicable(procedures, "above_mat");
            var roundApplicable = IsApplicable(procedures, "round");
            var duplicateApplicable = IsApplicable(procedures, "duplicate");
            var unbalancedApplicable = IsApplicable(procedures, "unbalanced");
            var lowFsliApplicable = IsApplicable(procedures, "low_fsli");
            var completenessApplicable = IsApplicable(procedures, "completeness");
            var benfordApplicable = IsApplicable(procedures, "benford");

            var holidayScope = procedureScopes.TryGetValue("holiday", out var holidayScopeValue)
                ? holidayScopeValue
                : null;

            var holidayVerification = master
                .Where(r => r.TestHolidayInfo == 1 && r.CheckDate.HasValue && !string.IsNullOrWhiteSpace(r.HolidayName))
                .GroupBy(r => new
                {
                    Date = r.CheckDate!.Value.Date,
                    Name = r.HolidayName.Trim(),
                })
                .OrderBy(g => g.Key.Date)
                .ThenBy(g => g.Key.Name, StringComparer.OrdinalIgnoreCase)
                .Select(g =>
                {
                    var scopedRows = g.Where(r => InScope(r, holidayScope)).ToList();
                    var scopedExceptionRows = scopedRows.Where(r => r.TestHoliday == 1).ToList();

                    return new HolidayVerificationRow
                    {
                        CheckDate = g.Key.Date,
                        HolidayName = g.Key.Name,
                        MatchedLines = g.Count(),
                        MatchedJournalGroups = g
                            .Select(r => r.JournalCompositeKey)
                            .Where(k => !string.IsNullOrWhiteSpace(k))
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .Count(),
                        InScopeLines = scopedRows.Count,
                        InScopeJournalGroups = scopedRows
                            .Select(r => r.JournalCompositeKey)
                            .Where(k => !string.IsNullOrWhiteSpace(k))
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .Count(),
                        InScopeExceptionLines = scopedExceptionRows.Count,
                        InScopeExceptionJournalGroups = scopedExceptionRows
                            .Select(r => r.JournalCompositeKey)
                            .Where(k => !string.IsNullOrWhiteSpace(k))
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .Count(),
                    };
                })
                .ToList();

            var result = new RunTestsResponse
            {
                Success = true,
                Message = "All tests complete.",
                Counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
                {
                    ["Weekend Journals"] = CountScopedJournalGroups(weekendApplicable, "weekend", x => x.TestWeekend == 1),
                    ["Holiday Journals"] = CountScopedJournalGroups(holidayApplicable, "holiday", x => x.TestHoliday == 1),
                    ["Backdated Journals"] = CountScopedJournalGroups(backdatedApplicable, "backdated", x => x.TestBackdated == 1),
                    ["Adj/Correc Descriptions"] = CountScopedJournalGroups(adjApplicable, "adj_desc", x => x.TestAdjDesc == 1),
                    ["Journals Above Materiality"] = CountScopedJournalGroups(matApplicable, "above_mat", x => x.TestAboveMat == 1),
                    ["Round Amount Analysis"] = CountScopedJournalGroups(roundApplicable, "round", x => x.TestRound == 1),
                    ["Duplicate Entries"] = CountScopedJournalGroups(duplicateApplicable, "duplicate", x => x.TestDuplicate == 1),
                    ["Unbalanced Journals"] = CountScopedJournalGroups(unbalancedApplicable, "unbalanced", x => x.TestUnbalanced == 1),
                    ["Low FSLI"] = CountScopedJournalGroups(lowFsliApplicable, "low_fsli", x => x.TestLowFsli == 1),
                },
                ReconAgree = completenessApplicable ? recon.Count(x => x.Status.Contains("Agrees", StringComparison.Ordinal)) : -1,
                ReconVariance = completenessApplicable ? recon.Count(x => x.Status.Contains("Variance", StringComparison.Ordinal)) : -1,
                ReconTotalRows = recon.Count,
                TotalDebit = master.Sum(x => x.Debit),
                TotalCredit = master.Sum(x => x.Credit),
                BenfordFlaggedDigits = benfordApplicable ? benford.Count(x => x.Status.Contains("⚠", StringComparison.Ordinal)) : -1,
                Benford = benfordApplicable ? benford : [],
                ReconTop = recon.Take(50).ToList(),
                ReconRows = recon.Take(250).ToList(),
                Users = ["All", ..master.Select(x => x.User).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x, StringComparer.OrdinalIgnoreCase)],
                Accounts = ["All", ..master.Select(x => x.Acct).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x, StringComparer.OrdinalIgnoreCase)],
                JournalKeys = ["All", ..master.Select(x => x.JournalKey).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x, StringComparer.OrdinalIgnoreCase).Take(500)],
                TotalRiskScore = decimal.Round(totalRiskScore, 2),
                RiskScoreSummary = riskSummary,
                UserAnalysisFull = userAnalysisFull,
                ProcedureResultsRanked = procedureRanked,
                IriRankedSummary = iriRanked,
                IriDetailRows = iriDetails.Take(300).ToList(),
                HolidayProcedureApplicable = holidayApplicable,
                HolidayVerification = holidayVerification,
            };

            return Ok(result);
        }
        catch (Exception ex)
        {
            return BadRequest(new RunTestsResponse { Success = false, Message = ex.Message });
        }
    }

    [HttpPost("export")]
    public ActionResult<ExportResponse> Export([FromBody] ExportRequest req)
    {
        try
        {
            var res = _export.Export(_state, req);
            if (!res.Success)
            {
                return BadRequest(res);
            }

            return Ok(res);
        }
        catch (Exception ex)
        {
            return BadRequest(new ExportResponse { Success = false, Message = ex.Message });
        }
    }

    [HttpGet("download")]
    public IActionResult Download([FromQuery] string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return BadRequest("Missing file path.");
        }

        try
        {
            var fullPath = Path.GetFullPath(path);
            if (!System.IO.File.Exists(fullPath))
            {
                return NotFound("File not found.");
            }

            var allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                ".xlsx", ".docx", ".csv", ".zip"
            };
            var ext = Path.GetExtension(fullPath);
            if (!allowed.Contains(ext))
            {
                return BadRequest("Unsupported file type.");
            }

            var provider = new FileExtensionContentTypeProvider();
            if (!provider.TryGetContentType(fullPath, out var contentType))
            {
                contentType = "application/octet-stream";
            }

            var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            return File(stream, contentType, Path.GetFileName(fullPath));
        }
        catch (Exception ex)
        {
            return BadRequest(ex.Message);
        }
    }

    [HttpPost("explore")]
    public ActionResult<ExploreResponse> Explore([FromBody] ExploreRequest req)
    {
        if (_state.MasterRows.Count == 0)
        {
            return BadRequest(new ExploreResponse { Success = false, Message = "Run tests first." });
        }

        try
        {
            IEnumerable<MasterRow> data = _state.MasterRows;

            data = req.Filter switch
            {
                "Weekend (exceptions)" => FullJournalRows(_state.MasterRows, x => x.TestWeekend == 1),
                "Weekend (all)" => FullJournalRows(_state.MasterRows, x => x.TestWeekendInfo == 1),
                "Holiday (exceptions)" => FullJournalRows(_state.MasterRows, x => x.TestHoliday == 1),
                "Holiday (all)" => FullJournalRows(_state.MasterRows, x => x.TestHolidayInfo == 1),
                "Backdated" or "Backdated Journals" => FullJournalRows(_state.MasterRows, x => x.TestBackdated == 1),
                "Adj Desc" or "Adjustment/Correction/Reversal Descriptions" => FullJournalRows(_state.MasterRows, x => x.TestAdjDesc == 1),
                "Above Mat." or "Journals Above Materiality" => FullJournalRows(_state.MasterRows, x => x.TestAboveMat == 1),
                "Round Amounts" or "Round Amount Analysis" => FullJournalRows(_state.MasterRows, x => x.TestRound == 1),
                "Duplicates" or "Duplicate Entries" => _state.MasterRows.Where(x => x.TestDuplicate == 1),
                "Unbalanced" or "Unbalanced Journals" => FullJournalRowsByGroupKey(_state.MasterRows, x => x.TestUnbalanced == 1),
                "Low FSLI" => FullJournalRows(_state.MasterRows, x => x.TestLowFsli == 1),
                "Manual Only" => _state.MasterRows.Where(x => x.IsManual == 1),
                "Automated Only" => _state.MasterRows.Where(x => x.IsManual == 0),
                _ => _state.MasterRows,
            };

            if (!string.Equals(req.User, "All", StringComparison.OrdinalIgnoreCase))
            {
                data = data.Where(x => string.Equals(x.User, req.User, StringComparison.OrdinalIgnoreCase));
            }

            if (!string.Equals(req.Account, "All", StringComparison.OrdinalIgnoreCase))
            {
                data = data.Where(x => string.Equals(x.Acct, req.Account, StringComparison.OrdinalIgnoreCase));
            }

            if (!string.Equals(req.JournalKey, "All", StringComparison.OrdinalIgnoreCase))
            {
                data = data.Where(x => string.Equals(x.JournalKey, req.JournalKey, StringComparison.OrdinalIgnoreCase));
            }

            var rows = data.ToList();
            var maxRows = req.MaxRows < 10 ? 10 : req.MaxRows;
            if (maxRows > 2000)
            {
                maxRows = 2000;
            }

            return Ok(new ExploreResponse
            {
                Success = true,
                Message = "Explorer results loaded.",
                Journals = rows.Select(x => x.JournalKey).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
                Lines = rows.Count,
                Rows = rows.Take(maxRows).ToList(),
            });
        }
        catch (Exception ex)
        {
            return BadRequest(new ExploreResponse { Success = false, Message = ex.Message });
        }
    }

    private static IEnumerable<MasterRow> FullJournalRows(IEnumerable<MasterRow> source, Func<MasterRow, bool> predicate)
    {
        var rows = source.ToList();
        var composites = rows.Where(predicate)
            .Select(x => x.JournalCompositeKey)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        return rows.Where(x => composites.Contains(x.JournalCompositeKey));
    }

    private static IEnumerable<MasterRow> FullJournalRowsByGroupKey(IEnumerable<MasterRow> source, Func<MasterRow, bool> predicate)
    {
        var rows = source.ToList();
        var keys = rows.Where(predicate)
            .Select(x => x.JournalKey)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        return rows.Where(x => keys.Contains(x.JournalKey));
    }

    private static string? ResolveDateColumn(DataTable gl, EngagementSettings eng, Dictionary<string, string> glMap)
    {
        if (glMap.TryGetValue("posting_date", out var postingDateCol) && !string.IsNullOrWhiteSpace(postingDateCol) && gl.Columns.Contains(postingDateCol))
        {
            return postingDateCol;
        }

        var preferred = new[] { "Posted_dt", "Doc_dt", "Posting_Date", "Creation_Date", "CPUDT" };
        foreach (var col in preferred)
        {
            if (gl.Columns.Contains(col))
            {
                return col;
            }
        }

        return gl.Columns.Cast<DataColumn>()
            .FirstOrDefault(c => c.ColumnName.Contains("date", StringComparison.OrdinalIgnoreCase))?.ColumnName;
    }

    private static bool HasColumnLike(DataTable table, string? mappedColumn, params string[] nameHints)
    {
        if (!string.IsNullOrWhiteSpace(mappedColumn) && table.Columns.Contains(mappedColumn))
        {
            return true;
        }

        var names = table.Columns.Cast<DataColumn>().Select(c => c.ColumnName);
        return names.Any(n => nameHints.Any(h => n.Contains(h, StringComparison.OrdinalIgnoreCase)));
    }

    private static bool TryDate(object? value, out DateTime dt)
    {
        dt = default;
        var parsed = CaatsAnalysisService.ParseFlexibleDate(value);
        if (!parsed.HasValue)
        {
            return false;
        }

        dt = parsed.Value;
        return true;
    }

    private static List<RiskScoreSummaryRow> BuildRiskScoreSummary(List<MasterRow> master)
    {
        const decimal rsManual = 3.0m;
        const decimal rsWeekend = 2.0m;
        const decimal rsHoliday = 2.0m;
        const decimal rsAdjDesc = 2.0m;
        const decimal rsBackdated = 3.0m;
        const decimal rsAboveMat = 3.0m;
        const decimal rsLowFsli = 2.0m;
        const decimal rsRound = 2.0m;
        const decimal rsUnbalanced = 3.0m;
        const decimal rsDuplicate = 2.0m;

        var rows = new List<RiskScoreSummaryRow>
        {
            BuildRiskRow("Manual Journals", master.Count(x => x.IsManual == 1), rsManual),
            BuildRiskRow("Weekend", master.Count(x => x.TestWeekend == 1), rsWeekend),
            BuildRiskRow("Holiday", master.Count(x => x.TestHoliday == 1), rsHoliday),
            BuildRiskRow("Adj Descriptions", master.Count(x => x.TestAdjDesc == 1), rsAdjDesc),
            BuildRiskRow("Backdated", master.Count(x => x.TestBackdated == 1), rsBackdated),
            BuildRiskRow("Above Materiality", master.Count(x => x.TestAboveMat == 1), rsAboveMat),
            BuildRiskRow("Low FSLI", master.Count(x => x.TestLowFsli == 1), rsLowFsli),
            BuildRiskRow("Round Amounts", master.Count(x => x.TestRound == 1), rsRound),
            BuildRiskRow("Unbalanced", master.Count(x => x.TestUnbalanced == 1), rsUnbalanced),
            BuildRiskRow("Duplicates", master.Count(x => x.TestDuplicate == 1), rsDuplicate),
        };

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

    private static List<UserAnalysisFullRow> BuildUserAnalysisFull(List<MasterRow> master)
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

    private static List<ProcedureRankedRow> BuildProcedureRanked(List<MasterRow> master, List<ReconRow> recon)
    {
        var items = new List<(string Procedure, int Count)>
        {
            ("Completeness", recon.Count(x => x.Status.Contains("Variance", StringComparison.OrdinalIgnoreCase))),
            ("Manual Weekend", master.Count(x => x.TestWeekend == 1)),
            ("Manual Holiday", master.Count(x => x.TestHoliday == 1)),
            ("User Analysis", master.Select(x => x.User).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).Count()),
            ("Unbalanced Journals", master.Count(x => x.TestUnbalanced == 1)),
            ("Above Performance Materiality", master.Count(x => x.TestAboveMat == 1)),
            ("Infrequent Financial Statement Line Item (FSLI)", master.Count(x => x.TestLowFsli == 1)),
            ("Manual Round Amounts", master.Count(x => x.TestRound == 1)),
            ("Backdated Entries", master.Count(x => x.TestBackdated == 1)),
            ("Description Contains Adjustment", master.Count(x => x.TestAdjDesc == 1)),
            ("Duplicates Entries", master.Count(x => x.TestDuplicate == 1)),
        };

        var ranked = items
            .OrderByDescending(x => x.Count)
            .ThenBy(x => x.Procedure, StringComparer.OrdinalIgnoreCase)
            .Select((x, idx) => new ProcedureRankedRow
            {
                Rank = idx + 1,
                Procedure = x.Procedure,
                Count = x.Count,
                RankedLabel = $"{idx + 1}: {x.Procedure}",
            })
            .ToList();

        return ranked;
    }

    private static List<IriRankedRow> BuildIriRanked(List<MasterRow> master, EngagementSettings engagement)
    {
        var detail = BuildIriDetails(master, engagement);
        var counts = detail.GroupBy(x => new { x.Iri, x.Description })
            .ToDictionary(g => g.Key, g => g.Count());

        var desc = IriDescriptions();
        var rows = desc
            .Select(kv => new IriRankedRow
            {
                Iri = kv.Key,
                Description = kv.Value,
                Count = counts.TryGetValue(new { Iri = kv.Key, Description = kv.Value }, out var c) ? c : 0,
            })
            .OrderByDescending(x => x.Count)
            .ThenBy(x => x.Iri, StringComparer.OrdinalIgnoreCase)
            .Select((x, idx) =>
            {
                x.Rank = idx + 1;
                x.RankedLabel = $"{idx + 1}: {x.Description}";
                return x;
            })
            .ToList();

        return rows;
    }

    private static List<IriDetailRow> BuildIriDetails(List<MasterRow> master, EngagementSettings engagement)
    {
        var desc = IriDescriptions();
        var outRows = new List<IriDetailRow>(capacity: master.Count / 2);
        var fyEnd = engagement.FyEnd?.Date;

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
                outRows.Add(ToIriDetail("SNG_IRI_01", desc["SNG_IRI_01"], row));
            }

            if (isManual && isBackdated && isRound && isReversal && isAbove)
            {
                outRows.Add(ToIriDetail("SNG_IRI_02", desc["SNG_IRI_02"], row));
            }

            if (isManual && isBackdated && isRound && isAbove)
            {
                outRows.Add(ToIriDetail("SNG_IRI_03", desc["SNG_IRI_03"], row));
            }

            if (isManual && isLowFsli && isReversal)
            {
                outRows.Add(ToIriDetail("SNG_IRI_04", desc["SNG_IRI_04"], row));
            }

            if (isManual && isLowFsli && isAbove)
            {
                outRows.Add(ToIriDetail("SNG_IRI_05", desc["SNG_IRI_05"], row));
            }

            if (fyEnd.HasValue && row.CpuDate.HasValue && row.CpuDate.Value.Date > fyEnd.Value)
            {
                outRows.Add(ToIriDetail("SNG_IRI_10", desc["SNG_IRI_10"], row));
            }
        }

        return outRows;
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
}
