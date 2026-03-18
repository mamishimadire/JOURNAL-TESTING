using System.Data;
using CaatsWebApp.Models.Caats;
using CaatsWebApp.Services.Caats;
using Xunit;

namespace CaatsWebApp.Tests;

public class CaatsAnalysisServiceTests
{
    [Fact]
    public void Run_ShouldUseJnlOriginOnly_WhenSngRuleIsOn()
    {
        var gl = BuildClassificationGlData();
        var tb = BuildClassificationTbData();
        var glMap = BuildClassificationGlMap();
        var tbMap = BuildClassificationTbMap();
        var settings = new EngagementSettings
        {
            SngRule = true,
            SngOriginsRaw = "FAJ,GJ,IJ,Journal,OBJ,PYRJ",
            PeriodStart = new DateTime(2024, 1, 1),
            PeriodEnd = new DateTime(2024, 12, 31),
        };

        var service = new CaatsAnalysisService();
        var (master, _, _) = service.Run(gl, tb, glMap, tbMap, settings, "GL", "DB");

        var manual = Assert.Single(master, x => x.JnlOrigin == "FAJ");
        var automated = Assert.Single(master, x => x.JnlOrigin == "AR");

        Assert.Equal(1, manual.IsManual);
        Assert.Equal("SNG:FAJ", manual.ManualReason);
        Assert.Equal(0, automated.IsManual);
        Assert.Equal("SNG-Auto:AR", automated.ManualReason);
    }

    [Fact]
    public void Run_ShouldClassifyByAutoIndicator_WhenSngRuleIsOff()
    {
        var gl = BuildClassificationGlData();
        var tb = BuildClassificationTbData();
        var glMap = BuildClassificationGlMap();
        var tbMap = BuildClassificationTbMap();
        var settings = new EngagementSettings
        {
            SngRule = false,
            PeriodStart = new DateTime(2024, 1, 1),
            PeriodEnd = new DateTime(2024, 12, 31),
        };

        var service = new CaatsAnalysisService();
        var (master, _, _) = service.Run(gl, tb, glMap, tbMap, settings, "GL", "DB");

        var manual = Assert.Single(master, x => x.JournalId == "DOC-MAN");
        var automated = Assert.Single(master, x => x.JournalId == "DOC-AUTO");

        Assert.Equal(1, manual.IsManual);
        Assert.StartsWith("STD:", manual.ManualReason, StringComparison.Ordinal);
        Assert.Equal(0, automated.IsManual);
        Assert.StartsWith("STD-Auto:", automated.ManualReason, StringComparison.Ordinal);
    }

    [Fact]
    public void Run_ShouldFlagCoreProcedures_WhenInputContainsKnownPatterns()
    {
        var gl = BuildGlData();
        var tb = BuildTbData();
        var glMap = BuildGlMap();
        var tbMap = BuildTbMap();
        var settings = BuildEngagement();

        var service = new CaatsAnalysisService();
        var (master, _, _) = service.Run(gl, tb, glMap, tbMap, settings, "GL", "DB");

        Assert.Contains(master, x => x.TestWeekendInfo == 1 && x.DayName.Equals("Saturday", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(master, x => x.TestHolidayInfo == 1 && x.HolidayName.Contains("Good Friday", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(master, x => x.TestBackdated == 1);
        Assert.Contains(master, x => x.TestAdjDesc == 1);
        Assert.Contains(master, x => x.TestAboveMat == 1);
        Assert.Contains(master, x => x.TestRound == 1 && x.RoundCategory.Contains("10,000", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(master, x => x.TestRound == 1 && x.RoundCategory.Contains("9,999", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(master, x => x.TestDuplicate == 1);
        Assert.Contains(master, x => x.TestUnbalanced == 1);
        Assert.Contains(master, x => x.TestLowFsli == 1);
    }

    [Fact]
    public void Run_ShouldProduceReconAndBenfordOutputs()
    {
        var gl = BuildGlData();
        var tb = BuildTbData();
        var glMap = BuildGlMap();
        var tbMap = BuildTbMap();
        var settings = BuildEngagement();

        var service = new CaatsAnalysisService();
        var (_, recon, benford) = service.Run(gl, tb, glMap, tbMap, settings, "GL", "DB");

        Assert.NotEmpty(recon);
        Assert.Contains(recon, x => x.Status.Contains("Agrees", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(recon, x => x.Status.Contains("Variance", StringComparison.OrdinalIgnoreCase));
        Assert.All(recon, x => Assert.Contains("Method: LAST value per Account_number", x.GlSource, StringComparison.OrdinalIgnoreCase));

        Assert.Equal(9, benford.Count);
        Assert.All(Enumerable.Range(1, 9), digit => Assert.Contains(benford, x => x.LeadingDigit == digit));
    }

    [Fact]
    public void Run_ShouldAnchorWeekendHolidayToPostingDate_AndApplyCoreFormulaFlags()
    {
        var gl = new DataTable();
        gl.Columns.Add("ACCOUNT", typeof(string));
        gl.Columns.Add("JOURNAL_KEY", typeof(string));
        gl.Columns.Add("JOURNAL_ID", typeof(string));
        gl.Columns.Add("POSTING_DATE", typeof(DateTime));
        gl.Columns.Add("CPU_DATE", typeof(DateTime));
        gl.Columns.Add("DESCRIPTION", typeof(string));
        gl.Columns.Add("USER_ID", typeof(string));
        gl.Columns.Add("PERIOD", typeof(string));
        gl.Columns.Add("JNL_ORIGIN", typeof(string));
        gl.Columns.Add("DEBIT", typeof(decimal));
        gl.Columns.Add("CREDIT", typeof(decimal));

        // Posting date weekday + CPU weekend: should NOT flag weekend.
        gl.Rows.Add("6000", "K10", "DOC-WKDAY", new DateTime(2024, 6, 14), new DateTime(2024, 6, 15), "normal", "u1", "2024-06", "FAJ", 100m, 0m);
        // Posting date holiday (Workers' Day): should flag holiday even if CPU is not holiday.
        gl.Rows.Add("6001", "K11", "DOC-HOL", new DateTime(2024, 5, 1), new DateTime(2024, 5, 2), "normal", "u2", "2024-05", "FAJ", 0m, 120m);
        // Adjustment keyword + above mat + round (10k .00)
        gl.Rows.Add("6002", "K12", "DOC-ROUND", new DateTime(2024, 6, 3), new DateTime(2024, 6, 3), "adjustment entry", "u3", "2024-06", "FAJ", 10_000m, 0m);
        // Round (9,999 any decimal)
        gl.Rows.Add("6003", "K13", "DOC-9999", new DateTime(2024, 6, 4), new DateTime(2024, 6, 4), "memo", "u4", "2024-06", "FAJ", 9_999.87m, 0m);

        var tb = new DataTable();
        tb.Columns.Add("ACCOUNT", typeof(string));
        tb.Columns.Add("NAME", typeof(string));
        tb.Columns.Add("OPENING", typeof(decimal));
        tb.Columns.Add("CLOSING", typeof(decimal));
        tb.Rows.Add("6000", "A", 0m, -100m);
        tb.Rows.Add("6001", "B", 0m, 120m);
        tb.Rows.Add("6002", "C", 0m, -10_000m);
        tb.Rows.Add("6003", "D", 0m, -9_999.87m);

        var glMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["account_number"] = "ACCOUNT",
            ["journal_id"] = "JOURNAL_ID",
            ["posting_date"] = "POSTING_DATE",
            ["creation_date"] = "CPU_DATE",
            ["description"] = "DESCRIPTION",
            ["user_id"] = "USER_ID",
            ["period"] = "PERIOD",
            ["jnl_origin"] = "JNL_ORIGIN",
            ["debit"] = "DEBIT",
            ["credit"] = "CREDIT",
        };
        var tbMap = BuildClassificationTbMap();

        var settings = new EngagementSettings
        {
            SngRule = true,
            SngOriginsRaw = "FAJ",
            CountryCode = "ZA",
            WeekendDateColumn = "CPU_DATE",
            WeekendNormal = false,
            HolidayNormal = false,
            PerformanceMateriality = 5_000m,
            PeriodStart = new DateTime(2024, 1, 1),
            PeriodEnd = new DateTime(2024, 12, 31),
        };

        var service = new CaatsAnalysisService();
        var (master, _, _) = service.Run(gl, tb, glMap, tbMap, settings, "GL", "DB");

        var weekdayPosting = Assert.Single(master, x => x.JournalId == "DOC-WKDAY");
        var holidayPosting = Assert.Single(master, x => x.JournalId == "DOC-HOL");
        var round10k = Assert.Single(master, x => x.JournalId == "DOC-ROUND");
        var round9999 = Assert.Single(master, x => x.JournalId == "DOC-9999");

        Assert.Equal(0, weekdayPosting.TestWeekend);
        Assert.Equal(0, weekdayPosting.TestWeekendInfo);
        Assert.Equal(1, holidayPosting.TestHoliday);
        Assert.Equal(1, holidayPosting.TestHolidayInfo);
        Assert.Equal(1, round10k.TestAdjDesc);
        Assert.Equal(1, round10k.TestAboveMat);
        Assert.Equal(1, round10k.TestRound);
        Assert.Equal("Multiple of 10,000 (.00)", round10k.RoundCategory);
        Assert.Equal(1, round9999.TestRound);
        Assert.Equal("9,999 Pattern (any decimal)", round9999.RoundCategory);
    }

    [Fact]
    public void Run_ShouldUseSouthAfricanHolidayCalendar_WithSpecialDaysAndSundayObservation()
    {
        var gl = new DataTable();
        gl.Columns.Add("ACCOUNT", typeof(string));
        gl.Columns.Add("JOURNAL_KEY", typeof(string));
        gl.Columns.Add("JOURNAL_ID", typeof(string));
        gl.Columns.Add("POSTING_DATE", typeof(DateTime));
        gl.Columns.Add("CPU_DATE", typeof(DateTime));
        gl.Columns.Add("DESCRIPTION", typeof(string));
        gl.Columns.Add("USER_ID", typeof(string));
        gl.Columns.Add("PERIOD", typeof(string));
        gl.Columns.Add("JNL_ORIGIN", typeof(string));
        gl.Columns.Add("DEBIT", typeof(decimal));
        gl.Columns.Add("CREDIT", typeof(decimal));

        // Sunday statutory holiday and observed Monday.
        gl.Rows.Add("7000", "H1", "DOC-NY-SUN", new DateTime(2023, 1, 1), new DateTime(2023, 1, 3), "new year sunday", "u1", "2023-01", "FAJ", 10m, 0m);
        gl.Rows.Add("7001", "H2", "DOC-NY-MON", new DateTime(2023, 1, 2), new DateTime(2023, 1, 3), "new year observed", "u1", "2023-01", "FAJ", 0m, 10m);

        // Special declared public holidays included in script/calendar set.
        gl.Rows.Add("7002", "H3", "DOC-RWC", new DateTime(2023, 12, 15), new DateTime(2023, 12, 15), "special holiday", "u2", "2023-12", "FAJ", 5m, 0m);
        gl.Rows.Add("7003", "H4", "DOC-ELECT", new DateTime(2024, 5, 29), new DateTime(2024, 5, 29), "election holiday", "u3", "2024-05", "FAJ", 0m, 5m);
        gl.Rows.Add("7004", "H5", "DOC-2022-SPECIAL", new DateTime(2022, 12, 27), new DateTime(2022, 12, 27), "goodwill displaced", "u4", "2022-12", "FAJ", 2m, 0m);

        // Non-holiday control date.
        gl.Rows.Add("7005", "H6", "DOC-NON-HOL", new DateTime(2024, 5, 30), new DateTime(2024, 5, 30), "non holiday", "u5", "2024-05", "FAJ", 1m, 0m);

        var tb = new DataTable();
        tb.Columns.Add("ACCOUNT", typeof(string));
        tb.Columns.Add("NAME", typeof(string));
        tb.Columns.Add("OPENING", typeof(decimal));
        tb.Columns.Add("CLOSING", typeof(decimal));
        tb.Rows.Add("7000", "A", 0m, -10m);
        tb.Rows.Add("7001", "B", 0m, 10m);
        tb.Rows.Add("7002", "C", 0m, -5m);
        tb.Rows.Add("7003", "D", 0m, 5m);
        tb.Rows.Add("7004", "E", 0m, -2m);
        tb.Rows.Add("7005", "F", 0m, -1m);

        var glMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["account_number"] = "ACCOUNT",
            ["journal_id"] = "JOURNAL_ID",
            ["posting_date"] = "POSTING_DATE",
            ["creation_date"] = "CPU_DATE",
            ["description"] = "DESCRIPTION",
            ["user_id"] = "USER_ID",
            ["period"] = "PERIOD",
            ["jnl_origin"] = "JNL_ORIGIN",
            ["debit"] = "DEBIT",
            ["credit"] = "CREDIT",
        };

        var settings = new EngagementSettings
        {
            SngRule = true,
            SngOriginsRaw = "FAJ",
            CountryCode = "ZA",
            WeekendNormal = false,
            HolidayNormal = false,
            PeriodStart = new DateTime(2022, 1, 1),
            PeriodEnd = new DateTime(2024, 12, 31),
        };

        var service = new CaatsAnalysisService();
        var (master, _, _) = service.Run(gl, tb, glMap, BuildClassificationTbMap(), settings, "GL", "DB");

        Assert.Equal(1, Assert.Single(master, x => x.JournalId == "DOC-NY-SUN").TestHolidayInfo);
        Assert.Equal(1, Assert.Single(master, x => x.JournalId == "DOC-NY-MON").TestHolidayInfo);
        Assert.Equal(1, Assert.Single(master, x => x.JournalId == "DOC-RWC").TestHolidayInfo);
        Assert.Equal(1, Assert.Single(master, x => x.JournalId == "DOC-ELECT").TestHolidayInfo);
        Assert.Equal(1, Assert.Single(master, x => x.JournalId == "DOC-2022-SPECIAL").TestHolidayInfo);
        Assert.Equal(0, Assert.Single(master, x => x.JournalId == "DOC-NON-HOL").TestHolidayInfo);
    }

    [Fact]
    public void Run_ShouldParsePostingDateDayFirst_ForHolidayWeekendFlags()
    {
        var gl = new DataTable();
        gl.Columns.Add("ACCOUNT", typeof(string));
        gl.Columns.Add("JOURNAL_KEY", typeof(string));
        gl.Columns.Add("JOURNAL_ID", typeof(string));
        gl.Columns.Add("POSTING_DATE", typeof(string));
        gl.Columns.Add("CPU_DATE", typeof(string));
        gl.Columns.Add("DESCRIPTION", typeof(string));
        gl.Columns.Add("USER_ID", typeof(string));
        gl.Columns.Add("PERIOD", typeof(string));
        gl.Columns.Add("JNL_ORIGIN", typeof(string));
        gl.Columns.Add("DEBIT", typeof(decimal));
        gl.Columns.Add("CREDIT", typeof(decimal));

        // dd/MM/yyyy format: must parse as 1 May 2024 (Workers' Day), not 5 Jan 2024.
        gl.Rows.Add("8000", "D1", "DOC-DAYFIRST-HOL", "01/05/2024", "02/05/2024", "holiday check", "u1", "2024-05", "FAJ", 100m, 0m);

        var tb = new DataTable();
        tb.Columns.Add("ACCOUNT", typeof(string));
        tb.Columns.Add("NAME", typeof(string));
        tb.Columns.Add("OPENING", typeof(decimal));
        tb.Columns.Add("CLOSING", typeof(decimal));
        tb.Rows.Add("8000", "A", 0m, -100m);

        var glMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["account_number"] = "ACCOUNT",
            ["journal_id"] = "JOURNAL_ID",
            ["posting_date"] = "POSTING_DATE",
            ["creation_date"] = "CPU_DATE",
            ["description"] = "DESCRIPTION",
            ["user_id"] = "USER_ID",
            ["period"] = "PERIOD",
            ["jnl_origin"] = "JNL_ORIGIN",
            ["debit"] = "DEBIT",
            ["credit"] = "CREDIT",
        };

        var settings = new EngagementSettings
        {
            SngRule = true,
            SngOriginsRaw = "FAJ",
            CountryCode = "ZA",
            WeekendNormal = false,
            HolidayNormal = false,
            PeriodStart = new DateTime(2024, 1, 1),
            PeriodEnd = new DateTime(2024, 12, 31),
        };

        var service = new CaatsAnalysisService();
        var (master, _, _) = service.Run(gl, tb, glMap, BuildClassificationTbMap(), settings, "GL", "DB");

        var row = Assert.Single(master, x => x.JournalId == "DOC-DAYFIRST-HOL");
        Assert.Equal(new DateTime(2024, 5, 1), row.PostingDate?.Date);
        Assert.Equal(1, row.TestHolidayInfo);
        Assert.Contains("Workers", row.HolidayName, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Run_ShouldHandleMultipleDateFormats_ForHolidayDetection()
    {
        var gl = new DataTable();
        gl.Columns.Add("ACCOUNT", typeof(string));
        gl.Columns.Add("JOURNAL_KEY", typeof(string));
        gl.Columns.Add("JOURNAL_ID", typeof(string));
        gl.Columns.Add("POSTING_DATE", typeof(string));
        gl.Columns.Add("CPU_DATE", typeof(string));
        gl.Columns.Add("DESCRIPTION", typeof(string));
        gl.Columns.Add("USER_ID", typeof(string));
        gl.Columns.Add("PERIOD", typeof(string));
        gl.Columns.Add("JNL_ORIGIN", typeof(string));
        gl.Columns.Add("DEBIT", typeof(decimal));
        gl.Columns.Add("CREDIT", typeof(decimal));

        var formats = new[]
        {
            "01/05/2024",                 // dd/MM/yyyy
            "1-5-2024",                   // d-M-yyyy
            "2024-05-01",                 // ISO
            "20240501",                   // yyyyMMdd
            "May 1, 2024",                // Month name
            "01 May 2024",                // dd MMM yyyy
            "2024-05-01T13:45:00",        // ISO datetime
            "1714521600",                 // Unix seconds
            "45413",                      // Excel/OA serial date
        };

        for (var i = 0; i < formats.Length; i++)
        {
            var acct = (8100 + i).ToString();
            gl.Rows.Add(acct, $"F{i}", $"DOC-FMT-{i}", formats[i], formats[i], "format holiday check", "u1", "2024-05", "FAJ", 100m, 0m);
        }

        var tb = new DataTable();
        tb.Columns.Add("ACCOUNT", typeof(string));
        tb.Columns.Add("NAME", typeof(string));
        tb.Columns.Add("OPENING", typeof(decimal));
        tb.Columns.Add("CLOSING", typeof(decimal));
        for (var i = 0; i < formats.Length; i++)
        {
            var acct = (8100 + i).ToString();
            tb.Rows.Add(acct, $"A{i}", 0m, -100m);
        }

        var glMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["account_number"] = "ACCOUNT",
            ["journal_id"] = "JOURNAL_ID",
            ["posting_date"] = "POSTING_DATE",
            ["creation_date"] = "CPU_DATE",
            ["description"] = "DESCRIPTION",
            ["user_id"] = "USER_ID",
            ["period"] = "PERIOD",
            ["jnl_origin"] = "JNL_ORIGIN",
            ["debit"] = "DEBIT",
            ["credit"] = "CREDIT",
        };

        var settings = new EngagementSettings
        {
            SngRule = true,
            SngOriginsRaw = "FAJ",
            CountryCode = "ZA",
            WeekendNormal = false,
            HolidayNormal = false,
            PeriodStart = new DateTime(2024, 1, 1),
            PeriodEnd = new DateTime(2024, 12, 31),
        };

        var service = new CaatsAnalysisService();
        var (master, _, _) = service.Run(gl, tb, glMap, BuildClassificationTbMap(), settings, "GL", "DB");

        Assert.Equal(formats.Length, master.Count);
        foreach (var row in master)
        {
            Assert.Equal(new DateTime(2024, 5, 1), row.PostingDate?.Date);
            Assert.Equal(1, row.TestHolidayInfo);
            Assert.Contains("Workers", row.HolidayName, StringComparison.OrdinalIgnoreCase);
        }
    }

    private static EngagementSettings BuildEngagement()
    {
        return new EngagementSettings
        {
            GlSystem = "Sage Intacct",
            CountryCode = "ZA",
            WeekendNormal = false,
            HolidayNormal = false,
            PerformanceMateriality = 5_000m,
            LowFsliThreshold = 2,
            SngRule = true,
            JournalGroupColumn = "JOURNAL_KEY",
            PeriodStart = new DateTime(2024, 1, 1),
            PeriodEnd = new DateTime(2024, 12, 31),
        };
    }

    private static Dictionary<string, string> BuildGlMap()
    {
        return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["account_number"] = "ACCOUNT",
            ["journal_id"] = "JOURNAL_ID",
            ["posting_date"] = "POSTING_DATE",
            ["creation_date"] = "CPU_DATE",
            ["description"] = "DESCRIPTION",
            ["user_id"] = "USER_ID",
            ["period"] = "PERIOD",
            ["auto_indicator"] = "AUTO_IND",
            ["dc_indicator"] = "DC_IND",
            ["jnl_origin"] = "JNL_ORIGIN",
            ["debit"] = "DEBIT",
            ["credit"] = "CREDIT",
            ["amount_signed"] = "AMOUNT_SIGNED",
        };
    }

    private static Dictionary<string, string> BuildTbMap()
    {
        return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["account_number"] = "ACCOUNT",
            ["account_name"] = "NAME",
            ["opening_bal"] = "OPENING",
            ["closing_bal"] = "CLOSING",
        };
    }

    private static DataTable BuildGlData()
    {
        var dt = new DataTable();
        dt.Columns.Add("ACCOUNT", typeof(string));
        dt.Columns.Add("JOURNAL_KEY", typeof(string));
        dt.Columns.Add("JOURNAL_ID", typeof(string));
        dt.Columns.Add("POSTING_DATE", typeof(DateTime));
        dt.Columns.Add("CPU_DATE", typeof(DateTime));
        dt.Columns.Add("DESCRIPTION", typeof(string));
        dt.Columns.Add("USER_ID", typeof(string));
        dt.Columns.Add("PERIOD", typeof(string));
        dt.Columns.Add("AUTO_IND", typeof(string));
        dt.Columns.Add("DC_IND", typeof(string));
        dt.Columns.Add("JNL_ORIGIN", typeof(string));
        dt.Columns.Add("DEBIT", typeof(decimal));
        dt.Columns.Add("CREDIT", typeof(decimal));
        dt.Columns.Add("AMOUNT_SIGNED", typeof(decimal));

        dt.Rows.Add("1000", "J1", "DOC-1", new DateTime(2024, 6, 15), new DateTime(2024, 6, 15), "adjust entry", "userA", "2024-06", "", "D", "FAJ", 100m, 0m, -100m);
        dt.Rows.Add("2000", "J1", "DOC-1", new DateTime(2024, 6, 15), new DateTime(2024, 6, 15), "adjust entry", "userA", "2024-06", "", "C", "FAJ", 0m, 90m, 90m);

        dt.Rows.Add("3000", "J2", "DOC-2", new DateTime(2024, 5, 1), new DateTime(2024, 6, 2), "period correction", "userB", "2024-05", "", "D", "GJ", 10_000m, 0m, -10_000m);

        dt.Rows.Add("4000", "J3", "DOC-3", new DateTime(2024, 4, 10), new DateTime(2024, 4, 10), "dup line", "userC", "2024-04", "", "D", "IJ", 9_999m, 0m, -9_999m);
        dt.Rows.Add("4000", "J3", "DOC-3", new DateTime(2024, 4, 10), new DateTime(2024, 4, 10), "dup line", "userC", "2024-04", "", "D", "IJ", 9_999m, 0m, -9_999m);

        dt.Rows.Add("5000", "J4", "DOC-4", new DateTime(2024, 3, 29), new DateTime(2024, 3, 29), "holiday post", "userD", "2024-03", "", "C", "OBJ", 0m, 50m, 50m);

        return dt;
    }

    private static DataTable BuildTbData()
    {
        var dt = new DataTable();
        dt.Columns.Add("ACCOUNT", typeof(string));
        dt.Columns.Add("NAME", typeof(string));
        dt.Columns.Add("OPENING", typeof(decimal));
        dt.Columns.Add("CLOSING", typeof(decimal));

        dt.Rows.Add("1000", "Cash", 0m, -100m);
        dt.Rows.Add("2000", "Revenue", 0m, 90m);
        dt.Rows.Add("9999", "Missing GL Account", 0m, 500m);

        return dt;
    }

    private static DataTable BuildClassificationGlData()
    {
        var dt = new DataTable();
        dt.Columns.Add("ACCOUNT", typeof(string));
        dt.Columns.Add("JOURNAL_KEY", typeof(string));
        dt.Columns.Add("JOURNAL_ID", typeof(string));
        dt.Columns.Add("POSTING_DATE", typeof(DateTime));
        dt.Columns.Add("CPU_DATE", typeof(DateTime));
        dt.Columns.Add("DESCRIPTION", typeof(string));
        dt.Columns.Add("USER_ID", typeof(string));
        dt.Columns.Add("PERIOD", typeof(string));
        dt.Columns.Add("AUTO_IND", typeof(string));
        dt.Columns.Add("JNL_ORIGIN", typeof(string));
        dt.Columns.Add("DEBIT", typeof(decimal));
        dt.Columns.Add("CREDIT", typeof(decimal));

        dt.Rows.Add("1000", "K1", "DOC-MAN", new DateTime(2024, 6, 3), new DateTime(2024, 6, 3), "manual", "u1", "2024-06", "", "FAJ", 10m, 0m);
        dt.Rows.Add("2000", "K2", "DOC-AUTO", new DateTime(2024, 6, 3), new DateTime(2024, 6, 3), "auto", "u2", "2024-06", "AUTO", "AR", 0m, 10m);

        return dt;
    }

    private static DataTable BuildClassificationTbData()
    {
        var dt = new DataTable();
        dt.Columns.Add("ACCOUNT", typeof(string));
        dt.Columns.Add("NAME", typeof(string));
        dt.Columns.Add("OPENING", typeof(decimal));
        dt.Columns.Add("CLOSING", typeof(decimal));
        dt.Rows.Add("1000", "Account 1000", 0m, -10m);
        dt.Rows.Add("2000", "Account 2000", 0m, 10m);
        return dt;
    }

    private static Dictionary<string, string> BuildClassificationGlMap()
    {
        return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["account_number"] = "ACCOUNT",
            ["journal_id"] = "JOURNAL_ID",
            ["posting_date"] = "POSTING_DATE",
            ["creation_date"] = "CPU_DATE",
            ["description"] = "DESCRIPTION",
            ["user_id"] = "USER_ID",
            ["period"] = "PERIOD",
            ["auto_indicator"] = "AUTO_IND",
            ["jnl_origin"] = "JNL_ORIGIN",
            ["debit"] = "DEBIT",
            ["credit"] = "CREDIT",
        };
    }

    private static Dictionary<string, string> BuildClassificationTbMap()
    {
        return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["account_number"] = "ACCOUNT",
            ["account_name"] = "NAME",
            ["opening_bal"] = "OPENING",
            ["closing_bal"] = "CLOSING",
        };
    }
}
