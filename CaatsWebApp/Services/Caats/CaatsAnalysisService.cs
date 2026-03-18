using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Text.Json;
using CaatsWebApp.Models.Caats;

namespace CaatsWebApp.Services.Caats;

public sealed class CaatsAnalysisService
{
    private static readonly Dictionary<int, decimal> BenfordExpected = new()
    {
        [1] = 30.1m, [2] = 17.6m, [3] = 12.5m, [4] = 9.7m, [5] = 7.9m,
        [6] = 6.7m, [7] = 5.8m, [8] = 5.1m, [9] = 4.6m,
    };

    public (List<MasterRow> Master, List<ReconRow> Recon, List<BenfordRow> Benford) Run(
        DataTable gl,
        DataTable tb,
        Dictionary<string, string> glMap,
        Dictionary<string, string> tbMap,
        EngagementSettings eng,
        string glTableName,
        string dbName)
    {
        var glRows = gl.AsEnumerable().Select(ToDictionary).ToList();
        var glRawRows = glRows.Select(r => new Dictionary<string, object?>(r, StringComparer.OrdinalIgnoreCase)).ToList();

        var master = BuildMasterRows(glRows, glMap, eng);
        master = ApplyDateFilter(master, eng);

        ApplyHolidayWeekendFlags(master, eng);
        ApplyPatternFlags(master, eng);
        ApplyDuplicateFlags(master);
        ApplyUnbalancedFlags(master);
        ApplyLowFsliFlags(master, eng.LowFsliThreshold);

        var recon = BuildRecon(glRawRows, tb.AsEnumerable().Select(ToDictionary).ToList(), glMap, tbMap, eng, glTableName, dbName);
        var benford = BenfordAnalysis(master.Select(x => x.AbsAmount));

        return (master, recon, benford);
    }

    private static List<MasterRow> BuildMasterRows(List<Dictionary<string, object?>> glRows, Dictionary<string, string> glMap, EngagementSettings eng)
    {
        var sngOrigins = (eng.SngOriginsRaw ?? string.Empty)
            .Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
            .Select(x => x.ToUpperInvariant())
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var rows = new List<MasterRow>(glRows.Count);
        foreach (var row in glRows)
        {
            var acctRaw = GetMapValue(row, glMap, "account_number")?.ToString();
            var acct = NormalizeAcct(acctRaw);
            var postDt = ParseFlexibleDate(GetMapValue(row, glMap, "posting_date"));
            var cpuDt = ParseFlexibleDate(GetMapValue(row, glMap, "creation_date")) ?? postDt;
            // Weekend and holiday checks are always anchored to posting/effective date.
            var wkndDt = postDt;

            var journalId = (GetMapValue(row, glMap, "journal_id")?.ToString() ?? string.Empty).Trim();
            var desc = (GetMapValue(row, glMap, "description")?.ToString() ?? string.Empty).Trim();
            var user = (GetMapValue(row, glMap, "user_id")?.ToString() ?? "Unknown").Trim();
            var period = (GetMapValue(row, glMap, "period")?.ToString() ?? string.Empty).Trim();
            var dc = (GetMapValue(row, glMap, "dc_indicator")?.ToString() ?? string.Empty).Trim().ToUpperInvariant();
            var origin = (GetMapValue(row, glMap, "jnl_origin")?.ToString() ?? string.Empty).Trim();
            var autoIndicatorRaw = (GetMapValue(row, glMap, "auto_indicator")?.ToString() ?? string.Empty).Trim();

            var grpKeyCol = eng.JournalGroupColumn;
            var grpKey = !string.IsNullOrWhiteSpace(grpKeyCol) && row.TryGetValue(grpKeyCol, out var grpRaw)
                ? (grpRaw?.ToString() ?? journalId).Trim()
                : journalId;

            decimal debit = 0;
            decimal credit = 0;
            decimal signed = 0;
            decimal abs = 0;

            var hasDr = glMap.TryGetValue("debit", out var drCol) && row.ContainsKey(drCol);
            var hasCr = glMap.TryGetValue("credit", out var crCol) && row.ContainsKey(crCol);
            var hasSigned = glMap.TryGetValue("amount_signed", out var signedCol) && row.ContainsKey(signedCol);
            var hasAbs = glMap.TryGetValue("abs_amount", out var absCol) && row.ContainsKey(absCol);

            if (hasDr && hasCr)
            {
                debit = Math.Abs(SafeDecimal(row[drCol!]));
                credit = Math.Abs(SafeDecimal(row[crCol!]));
                signed = credit - debit;
                abs = debit + credit;
            }
            else if (hasSigned)
            {
                signed = SafeDecimal(row[signedCol!]);
                abs = Math.Abs(signed);
                debit = signed < 0 ? abs : 0;
                credit = signed > 0 ? signed : 0;
            }
            else if (hasAbs)
            {
                abs = SafeDecimal(row[absCol!]);
                signed = dc is "S" or "D" ? -Math.Abs(abs) : Math.Abs(abs);
                debit = signed < 0 ? Math.Abs(signed) : 0;
                credit = signed > 0 ? signed : 0;
            }

            var postYm = postDt.HasValue ? postDt.Value.Year * 100 + postDt.Value.Month : 0;
            var cpuYm = cpuDt.HasValue ? cpuDt.Value.Year * 100 + cpuDt.Value.Month : 0;
            var backdatedMonths = Math.Max(cpuYm - postYm, 0);
            var backdatedDays = (cpuDt.HasValue && postDt.HasValue) ? Math.Max((cpuDt.Value - postDt.Value).Days, 0) : 0;

            // v15.1 parity:
            // SNG ON  => classify strictly by JNL Origin list (Auto Indicator ignored).
            // SNG OFF => classify by Auto Indicator (blank/0/contains MAN => manual).
            var isManual = false;
            var manualReason = string.Empty;

            if (eng.SngRule)
            {
                isManual = sngOrigins.Contains(origin.ToUpperInvariant());
                manualReason = isManual ? $"SNG:{origin}" : $"SNG-Auto:{origin}";
            }
            else
            {
                var autoUp = autoIndicatorRaw.ToUpperInvariant();
                isManual = string.IsNullOrWhiteSpace(autoUp) || autoUp == "0" || autoUp.Contains("MAN", StringComparison.Ordinal);
                manualReason = isManual
                    ? $"STD:{(string.IsNullOrWhiteSpace(autoIndicatorRaw) ? "BLANK" : autoIndicatorRaw)}"
                    : $"STD-Auto:{autoIndicatorRaw}";
            }

            var postingKey = postDt.HasValue ? postDt.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : string.Empty;
            var compositeKey = string.Concat(
                grpKey,
                "||",
                postingKey,
                "||",
                desc.ToLowerInvariant(),
                "||",
                debit.ToString("F2", CultureInfo.InvariantCulture),
                "||",
                credit.ToString("F2", CultureInfo.InvariantCulture));

            var (roundCategory, roundPattern) = GetRoundLabelDrCr(debit, credit);

            rows.Add(new MasterRow
            {
                Acct = acct,
                JournalKey = grpKey,
                JournalCompositeKey = compositeKey,
                JournalId = journalId,
                PostingDate = postDt,
                CpuDate = cpuDt,
                CheckDate = wkndDt,
                DayName = wkndDt?.DayOfWeek.ToString() ?? string.Empty,
                Description = desc,
                User = string.IsNullOrWhiteSpace(user) ? "Unknown" : user,
                Debit = debit,
                Credit = credit,
                Signed = signed,
                AbsAmount = abs,
                JnlOrigin = origin,
                IsManual = isManual ? 1 : 0,
                ManualReason = manualReason,
                MonthsBackdated = backdatedMonths,
                DaysBackdated = backdatedDays,
                RoundCategory = roundCategory ?? "Other",
                RoundPattern = roundPattern,
                BenfordDigit = GetLeadingDigit(abs),
                Period = period,
            });
        }

        return rows;
    }

    private static List<MasterRow> ApplyDateFilter(List<MasterRow> rows, EngagementSettings eng)
    {
        if (!eng.PeriodStart.HasValue || !eng.PeriodEnd.HasValue)
        {
            return rows;
        }

        var start = eng.PeriodStart.Value.Date;
        var end = eng.PeriodEnd.Value.Date;
        return rows.Where(r => r.PostingDate.HasValue && r.PostingDate.Value.Date >= start && r.PostingDate.Value.Date <= end).ToList();
    }

    private static void ApplyHolidayWeekendFlags(List<MasterRow> rows, EngagementSettings eng)
    {
        var years = rows.Where(x => x.CheckDate.HasValue).Select(x => x.CheckDate!.Value.Year).Distinct().ToList();
        var holidayMap = GetHolidays(eng.CountryCode, years);

        foreach (var row in rows)
        {
            var dt = row.CheckDate?.Date;
            var isWeekend = dt.HasValue && (dt.Value.DayOfWeek == DayOfWeek.Saturday || dt.Value.DayOfWeek == DayOfWeek.Sunday);
            var isHoliday = false;
            var holidayName = string.Empty;
            if (dt.HasValue && holidayMap.TryGetValue(dt.Value, out var hName))
            {
                isHoliday = true;
                holidayName = hName;
            }

            row.HolidayName = isHoliday ? holidayName : string.Empty;

            var manual = row.IsManual == 1;
            row.TestWeekendInfo = isWeekend ? 1 : 0;
            row.TestHolidayInfo = isHoliday ? 1 : 0;
            row.TestWeekendManual = isWeekend && manual ? 1 : 0;
            row.TestHolidayManual = isHoliday && manual ? 1 : 0;
            row.TestWeekend = isWeekend && !eng.WeekendNormal ? 1 : 0;
            row.TestHoliday = isHoliday && !eng.HolidayNormal ? 1 : 0;
            row.TestBackdated = row.MonthsBackdated > 0 ? 1 : 0;
            row.TestAdjDesc = Regex.IsMatch(row.Description, "adj|adjust|correc|revers|cancel|write.?off", RegexOptions.IgnoreCase) ? 1 : 0;
            row.TestAboveMat = row.AbsAmount >= eng.PerformanceMateriality ? 1 : 0;
            row.TestRound = row.RoundCategory != "Other" ? 1 : 0;
        }
    }

    private static void ApplyPatternFlags(List<MasterRow> rows, EngagementSettings eng)
    {
        _ = eng;
        foreach (var row in rows)
        {
            row.BenfordDigit = GetLeadingDigit(row.AbsAmount);
        }
    }

    private static void ApplyDuplicateFlags(List<MasterRow> rows)
    {
        var keys = rows.Select(r =>
                $"{r.JournalKey}|{r.Description.Trim().ToLowerInvariant()}|{r.PostingDate:yyyy-MM-dd}|{r.DayName}|{r.Debit:F2}|{r.Credit:F2}")
            .ToList();

        var dupSet = keys
            .GroupBy(x => x)
            .Where(g => g.Count() > 1)
            .Select(g => g.Key)
            .ToHashSet(StringComparer.Ordinal);

        for (var i = 0; i < rows.Count; i++)
        {
            rows[i].TestDuplicate = dupSet.Contains(keys[i]) ? 1 : 0;
        }
    }

    private static void ApplyUnbalancedFlags(List<MasterRow> rows)
    {
        var unbalanced = rows.GroupBy(r => r.JournalKey)
            .Where(g => Math.Abs(g.Sum(x => x.Debit) - g.Sum(x => x.Credit)) >= 1m)
            .Select(g => g.Key)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var row in rows)
        {
            row.TestUnbalanced = unbalanced.Contains(row.JournalKey) ? 1 : 0;
        }
    }

    private static void ApplyLowFsliFlags(List<MasterRow> rows, int lowFsliThreshold)
    {
        var lowSet = rows.GroupBy(r => new { r.Acct, r.Period })
            .Where(g => g.Count() < lowFsliThreshold)
            .Select(g => $"{g.Key.Acct}__{g.Key.Period}")
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var row in rows)
        {
            row.TestLowFsli = lowSet.Contains($"{row.Acct}__{row.Period}") ? 1 : 0;
        }
    }

    private static List<ReconRow> BuildRecon(
        List<Dictionary<string, object?>> glRawRows,
        List<Dictionary<string, object?>> tbRows,
        Dictionary<string, string> glMap,
        Dictionary<string, string> tbMap,
        EngagementSettings eng,
        string glTableName,
        string dbName)
    {
        var tbAcctCol = tbMap.GetValueOrDefault("account_number", string.Empty);
        var tbNameCol = tbMap.GetValueOrDefault("account_name", string.Empty);
        var tbOpenCol = tbMap.GetValueOrDefault("opening_bal", string.Empty);
        var tbCloseCol = tbMap.GetValueOrDefault("closing_bal", string.Empty);

        var tbDedup = tbRows
            .Select(r => new
            {
                Acct = NormalizeAcct(tbAcctCol.Length > 0 && r.TryGetValue(tbAcctCol, out var acctVal) ? acctVal?.ToString() : string.Empty),
                Name = tbNameCol.Length > 0 && r.TryGetValue(tbNameCol, out var nVal) ? (nVal?.ToString() ?? string.Empty).Trim() : string.Empty,
                Open = tbOpenCol.Length > 0 && r.TryGetValue(tbOpenCol, out var oVal) ? SafeDecimal(oVal) : 0m,
                Close = tbCloseCol.Length > 0 && r.TryGetValue(tbCloseCol, out var cVal) ? SafeDecimal(cVal) : 0m,
            })
            .Where(x => !string.IsNullOrWhiteSpace(x.Acct) && x.Acct != "0")
            .GroupBy(x => x.Acct)
            .Select(g => g.Last())
            .ToDictionary(x => x.Acct, x => x, StringComparer.OrdinalIgnoreCase);

        var glAcctCol = glMap.GetValueOrDefault("account_number", string.Empty);
        var glDrCol = glMap.GetValueOrDefault("debit", string.Empty);
        var glCrCol = glMap.GetValueOrDefault("credit", string.Empty);
        var glBalCol = string.IsNullOrWhiteSpace(eng.GlReconAmountColumn) ? glMap.GetValueOrDefault("amount_signed", string.Empty) : eng.GlReconAmountColumn;
        var hasReconBalanceColumn = !string.IsNullOrWhiteSpace(glBalCol) && glRawRows.Any(r => r.ContainsKey(glBalCol));

        var glValid = glRawRows
            .Select(r => new
            {
                Acct = NormalizeAcct(glAcctCol.Length > 0 && r.TryGetValue(glAcctCol, out var acctVal) ? acctVal?.ToString() : string.Empty),
                Debit = glDrCol.Length > 0 && r.TryGetValue(glDrCol, out var drVal) ? SafeDecimal(drVal) : 0m,
                Credit = glCrCol.Length > 0 && r.TryGetValue(glCrCol, out var crVal) ? SafeDecimal(crVal) : 0m,
                Balance = hasReconBalanceColumn && glBalCol.Length > 0 && r.TryGetValue(glBalCol, out var balVal) ? SafeDecimal(balVal) : 0m,
            })
            .Where(x => !string.IsNullOrWhiteSpace(x.Acct) && x.Acct != "0")
            .ToList();

        var glGrouped = glValid
            .GroupBy(x => x.Acct)
            .ToDictionary(
                g => g.Key,
                g => new
                {
                    GlDebit = g.Sum(x => Math.Abs(x.Debit)),
                    GlCredit = g.Sum(x => Math.Abs(x.Credit)),
                    GlBalance = hasReconBalanceColumn ? g.Last().Balance : g.Sum(x => Math.Abs(x.Credit)) - g.Sum(x => Math.Abs(x.Debit)),
                },
                StringComparer.OrdinalIgnoreCase);

        var allAccounts = tbDedup.Keys.Union(glGrouped.Keys, StringComparer.OrdinalIgnoreCase)
            .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var lookupColumn = hasReconBalanceColumn ? glBalCol : "Credit - Debit (fallback)";
        var glSource = $"DB: {dbName} | Table: {glTableName} | Col: {lookupColumn} | Method: LAST value per Account_number (gl_raw - pre-date-filter)";

        return allAccounts.Select(acct =>
        {
            var hasTb = tbDedup.TryGetValue(acct, out var tbRow);
            var hasGl = glGrouped.TryGetValue(acct, out var glRow);
            var open = hasTb ? tbRow!.Open : 0m;
            var close = hasTb ? tbRow!.Close : 0m;
            var dr = hasGl ? glRow!.GlDebit : 0m;
            var cr = hasGl ? glRow!.GlCredit : 0m;
            var bal = hasGl ? glRow!.GlBalance : 0m;
            var diff = decimal.Round(close - bal, 2);

            return new ReconRow
            {
                AccountNumber = acct,
                AccountName = hasTb ? tbRow!.Name : string.Empty,
                OpeningBalance = open,
                ClosingBalance = close,
                GlDebit = dr,
                GlCredit = cr,
                GlBalance = bal,
                Difference = diff,
                Status = Math.Abs(diff) < 1m ? "✅ Agrees" : "⚠️ Variance",
                GlSource = glSource,
            };
        }).ToList();
    }

    public static List<BenfordRow> BenfordAnalysis(IEnumerable<decimal> amounts)
    {
        var leadDigits = amounts.Where(x => x > 0m)
            .Select(GetLeadingDigit)
            .Where(x => x.HasValue)
            .Select(x => x!.Value)
            .ToList();

        if (leadDigits.Count == 0)
        {
            return [];
        }

        var total = leadDigits.Count;
        var rows = new List<BenfordRow>(9);
        for (var d = 1; d <= 9; d++)
        {
            var count = leadDigits.Count(x => x == d);
            var actual = total == 0 ? 0m : decimal.Round((decimal)count / total * 100m, 1);
            var expected = BenfordExpected[d];
            var diff = decimal.Round(actual - expected, 1);
            var status = diff switch
            {
                > 5m => "⚠️ Higher than expected",
                < -5m => "⚠️ Lower than expected",
                _ => "✅ Normal",
            };

            rows.Add(new BenfordRow
            {
                LeadingDigit = d,
                Count = count,
                ActualPercent = actual,
                ExpectedPercent = expected,
                DifferencePercent = diff,
                Status = status,
            });
        }

        return rows;
    }

    private static Dictionary<string, object?> ToDictionary(DataRow row)
    {
        return row.Table.Columns.Cast<DataColumn>()
            .ToDictionary(c => c.ColumnName, c => row[c] == DBNull.Value ? null : row[c], StringComparer.OrdinalIgnoreCase);
    }

    private static object? GetMapValue(Dictionary<string, object?> row, Dictionary<string, string> map, string key)
    {
        if (!map.TryGetValue(key, out var colName) || string.IsNullOrWhiteSpace(colName))
        {
            return null;
        }

        return row.TryGetValue(colName, out var val) ? val : null;
    }

    internal static DateTime? ParseFlexibleDate(object? value)
    {
        if (value is null)
        {
            return null;
        }

        if (value is DateTime dt)
        {
            return dt;
        }

        if (value is DateTimeOffset dto)
        {
            return dto.DateTime;
        }

        if (value is DateOnly dateOnly)
        {
            return dateOnly.ToDateTime(TimeOnly.MinValue);
        }

        if (TryParseNumericDateValue(value, out var parsedNumeric))
        {
            return parsedNumeric;
        }

        var raw = value.ToString()?.Trim();
        if (string.IsNullOrWhiteSpace(raw))
        {
            return null;
        }

        if (raw is "None" or "nan" or "NaT" or "NULL" or "null" or "N/A" or "-")
        {
            return null;
        }

        // Compact numeric dates often arrive as strings from CSV/Excel exports.
        if (TryParseNumericDateString(raw, out parsedNumeric))
        {
            return parsedNumeric;
        }

        if (TryParseDelimitedNumericDate(raw, out var parsedDelimited))
        {
            return parsedDelimited;
        }

        // Keep parity with notebook _safe_date: day-first first, then year-first fallback.
        var dayFirstFormats = new[]
        {
            "d/M/yyyy",
            "dd/MM/yyyy",
            "d-M-yyyy",
            "dd-MM-yyyy",
            "d.M.yyyy",
            "dd.MM.yyyy",
            "d/M/yy",
            "dd/MM/yy",
            "d-M-yy",
            "dd-MM-yy",
            "d/M/yyyy H:m",
            "dd/MM/yyyy HH:mm",
            "d/M/yyyy H:m:s",
            "dd/MM/yyyy HH:mm:ss",
            "d-M-yyyy H:m",
            "dd-MM-yyyy HH:mm",
            "d-M-yyyy H:m:s",
            "dd-MM-yyyy HH:mm:ss",
            "d/M/yyyy h:m tt",
            "dd/MM/yyyy hh:mm tt",
            "d-M-yyyy h:m tt",
            "dd-MM-yyyy hh:mm tt",
            "d.M.yyyy H:m",
            "dd.MM.yyyy HH:mm",
            "d.M.yyyy H:m:s",
            "dd.MM.yyyy HH:mm:ss",
        };

        if (DateTime.TryParseExact(raw, dayFirstFormats, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var parsedDayFirstExact))
        {
            return parsedDayFirstExact;
        }

        if (DateTime.TryParse(raw, new CultureInfo("en-ZA"), DateTimeStyles.AllowWhiteSpaces, out var parsedZa))
        {
            return parsedZa;
        }

        var monthNameFormats = new[]
        {
            "d MMM yyyy",
            "dd MMM yyyy",
            "d MMMM yyyy",
            "dd MMMM yyyy",
            "MMM d, yyyy",
            "MMMM d, yyyy",
            "d MMM yyyy H:m",
            "dd MMM yyyy HH:mm",
            "d MMMM yyyy H:m",
            "dd MMMM yyyy HH:mm",
            "MMM d, yyyy H:m",
            "MMMM d, yyyy HH:mm",
        };

        if (DateTime.TryParseExact(raw, monthNameFormats, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var parsedMonthName))
        {
            return parsedMonthName;
        }

        if (DateTimeOffset.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var parsedOffset))
        {
            return parsedOffset.DateTime;
        }

        if (DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var parsedInv))
        {
            return parsedInv;
        }

        if (DateTime.TryParse(raw, new CultureInfo("en-US"), DateTimeStyles.AllowWhiteSpaces, out var parsedUs))
        {
            return parsedUs;
        }

        var yearFirstFormats = new[]
        {
            "yyyy-MM-dd",
            "yyyy/MM/dd",
            "yyyy.MM.dd",
            "yyyyMMdd",
            "yyyy-MM-dd HH:mm:ss",
            "yyyy/MM/dd HH:mm:ss",
            "yyyy-MM-ddTHH:mm:ss",
            "yyyy-MM-ddTHH:mm:ss.fff",
            "yyyy-MM-ddTHH:mm:ssZ",
            "yyyy-MM-ddTHH:mm:ss.fffZ",
            "yyyy-MM-ddTHH:mm:ss.fffffff",
            "yyyy-MM-ddTHH:mm:ss.fffffffZ",
        };

        if (DateTime.TryParseExact(raw, yearFirstFormats, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var parsedYearFirst))
        {
            return parsedYearFirst;
        }

        return null;
    }

    private static bool TryParseNumericDateValue(object value, out DateTime parsed)
    {
        parsed = default;

        if (value is long l)
        {
            return TryParseNumericDateString(l.ToString(CultureInfo.InvariantCulture), out parsed);
        }

        if (value is int i)
        {
            return TryParseNumericDateString(i.ToString(CultureInfo.InvariantCulture), out parsed);
        }

        if (value is decimal dec)
        {
            return TryParseNumericDateString(dec.ToString(CultureInfo.InvariantCulture), out parsed);
        }

        if (value is double dbl)
        {
            return TryParseNumericDateString(dbl.ToString(CultureInfo.InvariantCulture), out parsed);
        }

        return false;
    }

    private static bool TryParseNumericDateString(string raw, out DateTime parsed)
    {
        parsed = default;
        var trimmed = raw.Trim();

        if (!Regex.IsMatch(trimmed, "^[0-9]+(?:\\.[0-9]+)?$", RegexOptions.CultureInvariant))
        {
            return false;
        }

        if (long.TryParse(trimmed, NumberStyles.Integer, CultureInfo.InvariantCulture, out var whole))
        {
            if (trimmed.Length == 8)
            {
                // yyyyMMdd first, then ddMMyyyy fallback for day-first parity.
                if (DateTime.TryParseExact(trimmed, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed))
                {
                    return true;
                }

                if (DateTime.TryParseExact(trimmed, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed))
                {
                    return true;
                }
            }

            if (trimmed.Length == 10)
            {
                try
                {
                    parsed = DateTimeOffset.FromUnixTimeSeconds(whole).DateTime;
                    return true;
                }
                catch
                {
                    // not a unix seconds value
                }
            }

            if (trimmed.Length == 13)
            {
                try
                {
                    parsed = DateTimeOffset.FromUnixTimeMilliseconds(whole).DateTime;
                    return true;
                }
                catch
                {
                    // not a unix milliseconds value
                }
            }
        }

        if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out var numeric)
            && numeric >= 1d && numeric <= 2_958_465d)
        {
            // Excel / OLE Automation serial date.
            try
            {
                parsed = DateTime.FromOADate(numeric);
                return true;
            }
            catch
            {
                // not a valid OA date
            }
        }

        return false;
    }

    private static bool TryParseDelimitedNumericDate(string raw, out DateTime parsed)
    {
        parsed = default;

        var datePart = raw;
        var timePart = string.Empty;

        var tIndex = raw.IndexOf('T', StringComparison.Ordinal);
        if (tIndex > 0)
        {
            datePart = raw[..tIndex];
            timePart = raw[(tIndex + 1)..];
        }
        else
        {
            var spaceIndex = raw.IndexOf(' ');
            if (spaceIndex > 0)
            {
                datePart = raw[..spaceIndex];
                timePart = raw[(spaceIndex + 1)..];
            }
        }

        if (!Regex.IsMatch(datePart, "^[0-9]{1,4}[-/.][0-9]{1,2}[-/.][0-9]{1,4}$", RegexOptions.CultureInvariant))
        {
            return false;
        }

        var parts = Regex.Split(datePart, "[-/.]");
        if (parts.Length != 3
            || !int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out var p1)
            || !int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var p2)
            || !int.TryParse(parts[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out var p3))
        {
            return false;
        }

        int year;
        int month;
        int day;

        if (parts[0].Length == 4)
        {
            year = p1;
            month = p2;
            day = p3;
        }
        else if (parts[2].Length == 4)
        {
            year = p3;
            if (p1 > 12 && p2 <= 12)
            {
                day = p1;
                month = p2;
            }
            else if (p2 > 12 && p1 <= 12)
            {
                month = p1;
                day = p2;
            }
            else
            {
                // Ambiguous values (both <= 12): default to day-first.
                day = p1;
                month = p2;
            }
        }
        else
        {
            return false;
        }

        try
        {
            parsed = new DateTime(year, month, day, 0, 0, 0, DateTimeKind.Unspecified);
        }
        catch
        {
            return false;
        }

        if (string.IsNullOrWhiteSpace(timePart))
        {
            return true;
        }

        if (DateTime.TryParse(timePart.Trim(), CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var parsedTime)
            || DateTime.TryParse(timePart.Trim(), new CultureInfo("en-ZA"), DateTimeStyles.AllowWhiteSpaces, out parsedTime)
            || DateTime.TryParse(timePart.Trim(), new CultureInfo("en-US"), DateTimeStyles.AllowWhiteSpaces, out parsedTime))
        {
            parsed = parsed.Date + parsedTime.TimeOfDay;
            return true;
        }

        return true;
    }

    private static string NormalizeAcct(string? val)
    {
        var s = (val ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(s) || s is "nan" or "None" or "NULL" or "null" or "NaN" or "NaT")
        {
            return string.Empty;
        }

        if (decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var num))
        {
            var asLong = (long)num;
            var normalized = asLong.ToString(CultureInfo.InvariantCulture).TrimStart('0');
            return normalized.Length == 0 ? "0" : normalized;
        }

        return s.ToUpperInvariant();
    }

    private static decimal SafeDecimal(object? value)
    {
        if (value is null)
        {
            return 0m;
        }

        var x = value.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(x) || x is "nan" or "None" or "NaT" or "NULL" or "null")
        {
            return 0m;
        }

        var negative = x.Contains('-', StringComparison.Ordinal);
        var clean = Regex.Replace(x, "[R$\\s+\\-]", string.Empty, RegexOptions.CultureInvariant);
        if (string.IsNullOrWhiteSpace(clean))
        {
            return 0m;
        }

        var nCommas = clean.Count(c => c == ',');
        var nDots = clean.Count(c => c == '.');

        if (nCommas == 0 && nDots <= 1)
        {
            // already clean
        }
        else if (nCommas == 1 && nDots == 0)
        {
            var afterComma = clean.Split(',')[1];
            clean = afterComma.Length == 3 ? clean.Replace(",", string.Empty, StringComparison.Ordinal) : clean.Replace(',', '.');
        }
        else if (nCommas == 0 && nDots > 1)
        {
            clean = clean.Replace(".", string.Empty, StringComparison.Ordinal);
        }
        else if (nCommas > 1 && nDots == 0)
        {
            clean = clean.Replace(",", string.Empty, StringComparison.Ordinal);
        }
        else
        {
            var lastComma = clean.LastIndexOf(",", StringComparison.Ordinal);
            var lastDot = clean.LastIndexOf(".", StringComparison.Ordinal);
            clean = lastDot > lastComma
                ? clean.Replace(",", string.Empty, StringComparison.Ordinal)
                : clean.Replace(".", string.Empty, StringComparison.Ordinal).Replace(',', '.');
        }

        if (!decimal.TryParse(clean, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed))
        {
            return 0m;
        }

        return negative ? -parsed : parsed;
    }

    private static (string? Category, string Label) RoundPattern(decimal amount)
    {
        var val = Math.Abs(amount);
        var whole = (long)Math.Truncate(val);
        var dec = (int)((val - whole) * 100m);

        if (val >= 10_000m && whole % 10_000 == 0 && dec == 0)
        {
            return ("Multiple of 10,000 (.00)", $"R{whole:N0}.00");
        }

        if (val >= 9_999m && whole % 10_000 == 9_999)
        {
            return ("9,999 Pattern (any decimal)", $"R{whole:N0}.{dec:00}");
        }

        return (null, "Other");
    }

    private static (string? Category, string Label) GetRoundLabelDrCr(decimal debit, decimal credit)
    {
        var dr = RoundPattern(debit);
        if (dr.Category is not null)
        {
            return dr;
        }

        var cr = RoundPattern(credit);
        return cr.Category is not null ? cr : (null, "Other");
    }

    private static int? GetLeadingDigit(decimal amount)
    {
        var val = Math.Abs(amount);
        if (val == 0m)
        {
            return null;
        }

        var s = val.ToString("0.################", CultureInfo.InvariantCulture)
            .Replace(".", string.Empty, StringComparison.Ordinal)
            .TrimStart('0');

        if (s.Length == 0 || !char.IsDigit(s[0]))
        {
            return null;
        }

        return s[0] - '0';
    }

    private static Dictionary<DateTime, string> GetHolidays(string countryCode, IEnumerable<int> years)
    {
        var map = new Dictionary<DateTime, string>();
        if (!string.Equals(countryCode, "ZA", StringComparison.OrdinalIgnoreCase))
        {
            using var http = new HttpClient();
            foreach (var y in years)
            {
                try
                {
                    var url = $"https://date.nager.at/api/v3/publicholidays/{y}/{countryCode}";
                    var json = http.GetStringAsync(url).GetAwaiter().GetResult();
                    var doc = JsonDocument.Parse(json);
                    foreach (var item in doc.RootElement.EnumerateArray())
                    {
                        var dateRaw = item.GetProperty("date").GetString();
                        if (!DateTime.TryParse(dateRaw, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var dt))
                        {
                            continue;
                        }

                        var name = item.TryGetProperty("name", out var nameProp)
                            ? nameProp.GetString()
                            : item.TryGetProperty("localName", out var localProp)
                                ? localProp.GetString()
                                : "Public Holiday";

                        map[dt.Date] = string.IsNullOrWhiteSpace(name) ? "Public Holiday" : name!;
                    }
                }
                catch
                {
                    // Keep parity with notebook behavior: if API fails, continue without throwing.
                }
            }

            return map;
        }

        foreach (var y in years)
        {
            foreach (var kv in SaFixedHolidays(y))
            {
                map[kv.Key] = kv.Value;
            }
        }

        return map;
    }

    private static Dictionary<DateTime, string> SaFixedHolidays(int year)
    {
        var easter = EasterSunday(year);
        var raw = new Dictionary<DateTime, string>
        {
            [new DateTime(year, 1, 1)] = "New Year's Day",
            [new DateTime(year, 3, 21)] = "Human Rights Day",
            [easter.AddDays(-2)] = "Good Friday",
            [easter.AddDays(1)] = "Family Day (Easter Monday)",
            [new DateTime(year, 4, 27)] = "Freedom Day",
            [new DateTime(year, 5, 1)] = "Workers' Day",
            [new DateTime(year, 6, 16)] = "Youth Day",
            [new DateTime(year, 8, 9)] = "National Women's Day",
            [new DateTime(year, 9, 24)] = "Heritage Day",
            [new DateTime(year, 12, 16)] = "Day of Reconciliation",
            [new DateTime(year, 12, 25)] = "Christmas Day",
            [new DateTime(year, 12, 26)] = "Day of Goodwill",
        };

        var result = new Dictionary<DateTime, string>();
        foreach (var kv in raw)
        {
            var adjusted = kv.Key.DayOfWeek == DayOfWeek.Sunday ? kv.Key.AddDays(1) : kv.Key;
            result[adjusted] = kv.Value;
            if (adjusted != kv.Key)
            {
                result[kv.Key] = kv.Value + " (observed)";
            }
        }

        var extras = new Dictionary<DateTime, string>
        {
            [new DateTime(2022, 12, 27)] = "Day of Goodwill (observed - Christmas displaced to 26 Dec)",
            [new DateTime(2023, 12, 15)] = "Special Public Holiday (Rugby World Cup)",
            [new DateTime(2024, 5, 29)] = "National Election Day",
        };

        foreach (var kv in extras.Where(x => x.Key.Year == year))
        {
            result[kv.Key] = kv.Value;
        }

        return result;
    }

    private static DateTime EasterSunday(int year)
    {
        var a = year % 19;
        var b = year / 100;
        var c = year % 100;
        var d = b / 4;
        var e = b % 4;
        var f = (b + 8) / 25;
        var g = (b - f + 1) / 3;
        var h = (19 * a + b - d - g + 15) % 30;
        var i = c / 4;
        var k = c % 4;
        var l = (32 + 2 * e + 2 * i - h - k) % 7;
        var m = (a + 11 * h + 22 * l) / 451;
        var month = (h + l - 7 * m + 114) / 31;
        var day = ((h + l - 7 * m + 114) % 31) + 1;
        return new DateTime(year, month, day);
    }
}
