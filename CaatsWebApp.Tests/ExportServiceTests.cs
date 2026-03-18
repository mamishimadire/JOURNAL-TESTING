using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using CaatsWebApp.Models.Caats;
using CaatsWebApp.Services.Caats;
using Xunit;

namespace CaatsWebApp.Tests;

public sealed class ExportServiceTests
{
    [Fact]
    public void Export_ShouldReturnFailure_WhenNoRunResultsFound()
    {
        var service = new ExportService();
        var state = new CaatsState();

        var result = service.Export(state, new ExportRequest
        {
            OutputFolder = Path.Combine(Path.GetTempPath(), "caats-empty-" + Guid.NewGuid().ToString("N")),
            Word = true,
            Excel = true,
            Csv = false,
        });

        Assert.False(result.Success);
        Assert.Contains("Run tests first", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Export_ShouldCreateWordExcelAndZip_WithExpectedContent()
    {
        var outputFolder = Path.Combine(Path.GetTempPath(), "caats-export-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(outputFolder);

        try
        {
            var service = new ExportService();
            var state = BuildState();

            var result = service.Export(state, new ExportRequest
            {
                OutputFolder = outputFolder,
                Word = true,
                Excel = true,
                Csv = false,
            });

            Assert.True(result.Success);
            var wordPath = Assert.Single(result.Files, x => x.EndsWith(".docx", StringComparison.OrdinalIgnoreCase));
            var excelPath = Assert.Single(result.Files, x => x.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase));
            var zipPath = Assert.Single(result.Files, x => x.EndsWith(".zip", StringComparison.OrdinalIgnoreCase));

            Assert.True(File.Exists(wordPath));
            Assert.True(File.Exists(excelPath));
            Assert.True(File.Exists(zipPath));

            using (var workbook = new XLWorkbook(excelPath))
            {
                Assert.NotNull(workbook.Worksheet("MASTER_Population"));
                Assert.NotNull(workbook.Worksheet("Holiday_Journals"));
                Assert.NotNull(workbook.Worksheet("Account_Analysis"));
                Assert.NotNull(workbook.Worksheet("GL_TB_Recon"));

                var wsMaster = workbook.Worksheet("MASTER_Population");
                Assert.Equal("Day Name", wsMaster.Cell(1, 6).GetString());
                Assert.Equal("Holiday Name", wsMaster.Cell(1, 7).GetString());

                var wsRecon = workbook.Worksheet("GL_TB_Recon");
                Assert.Equal("GL Source", wsRecon.Cell(1, 10).GetString());

                var wsWeekend = workbook.Worksheet("Weekend_Journals");
                var lastWeekendRow = wsWeekend.LastRowUsed();
                Assert.NotNull(lastWeekendRow);
                Assert.Equal(1, lastWeekendRow!.RowNumber());
            }

            using (var doc = WordprocessingDocument.Open(wordPath, false))
            {
                var body = doc.MainDocumentPart?.Document?.Body;
                Assert.NotNull(body);

                var text = body!.InnerText;
                Assert.Contains("Journal CAAT Objective", text, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("Analysis Steps and Audit Procedures", text, StringComparison.OrdinalIgnoreCase);

                var hasTocField = body.Descendants<FieldCode>()
                    .Any(x => x.InnerText.Contains("TOC", StringComparison.OrdinalIgnoreCase));
                Assert.True(hasTocField);
            }
        }
        finally
        {
            if (Directory.Exists(outputFolder))
            {
                Directory.Delete(outputFolder, true);
            }
        }
    }

    private static CaatsState BuildState()
    {
        return new CaatsState
        {
            Engagement = new EngagementSettings
            {
                Parent = "Parent",
                Client = "TestClient",
                Engagement = "JE Testing",
                GlSystem = "Sage Intacct",
                CountryCode = "ZA",
                WeekendNormal = false,
                HolidayNormal = false,
                PeriodStart = new DateTime(2024, 1, 1),
                PeriodEnd = new DateTime(2024, 12, 31),
                SigningDate = new DateTime(2026, 3, 6),
                Materiality = 10_000_000m,
                PerformanceMateriality = 7_500_000m,
                Auditor = "QA Auditor",
                Manager = "QA Manager",
            },
            MasterRows =
            [
                new MasterRow
                {
                    Acct = "1000",
                    JournalKey = "J1",
                    JournalId = "DOC1",
                    PostingDate = new DateTime(2024, 3, 29),
                    CpuDate = new DateTime(2024, 3, 29),
                    CheckDate = new DateTime(2024, 3, 29),
                    DayName = "Friday",
                    HolidayName = "Good Friday",
                    Description = "adjustment",
                    User = "userA",
                    Debit = 100m,
                    Credit = 0m,
                    Signed = -100m,
                    AbsAmount = 100m,
                    IsManual = 1,
                    ManualReason = "Indicator",
                    Period = "2024-03",
                    RoundCategory = "Other",
                    RoundPattern = "Other",
                    TestHolidayInfo = 1,
                    TestHoliday = 1,
                    TestAdjDesc = 1,
                },
                new MasterRow
                {
                    Acct = "2000",
                    JournalKey = "J1",
                    JournalId = "DOC1",
                    PostingDate = new DateTime(2024, 3, 30),
                    CpuDate = new DateTime(2024, 4, 2),
                    CheckDate = new DateTime(2024, 3, 30),
                    DayName = "Saturday",
                    Description = "normal entry",
                    User = "userB",
                    Debit = 0m,
                    Credit = 90m,
                    Signed = 90m,
                    AbsAmount = 90m,
                    IsManual = 1,
                    ManualReason = "Indicator",
                    Period = "2024-03",
                    RoundCategory = "9,999 Pattern (any decimal)",
                    RoundPattern = "R9,999.00",
                    MonthsBackdated = 1,
                    TestWeekendInfo = 1,
                    TestWeekend = 1,
                    TestRound = 1,
                    TestBackdated = 1,
                },
            ],
            ReconRows =
            [
                new ReconRow
                {
                    AccountNumber = "1000",
                    AccountName = "Cash",
                    OpeningBalance = 0m,
                    ClosingBalance = -100m,
                    GlDebit = 100m,
                    GlCredit = 0m,
                    GlBalance = -100m,
                    Difference = 0m,
                    Status = "✅ Agrees",
                    GlSource = "DB: DB | Table: GL | Col: Balance | Method: LAST value per Account_number (gl_raw - pre-date-filter)",
                },
                new ReconRow
                {
                    AccountNumber = "9999",
                    AccountName = "Suspense",
                    OpeningBalance = 0m,
                    ClosingBalance = 500m,
                    GlDebit = 0m,
                    GlCredit = 0m,
                    GlBalance = 0m,
                    Difference = 500m,
                    Status = "⚠️ Variance",
                    GlSource = "DB: DB | Table: GL | Col: Balance | Method: LAST value per Account_number (gl_raw - pre-date-filter)",
                },
            ],
            BenfordRows =
            [
                new BenfordRow { LeadingDigit = 1, Count = 1, ActualPercent = 50m, ExpectedPercent = 30.1m, DifferencePercent = 19.9m, Status = "⚠️ Higher than expected" },
                new BenfordRow { LeadingDigit = 9, Count = 1, ActualPercent = 50m, ExpectedPercent = 4.6m, DifferencePercent = 45.4m, Status = "⚠️ Higher than expected" },
            ],
            LastProcedures = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase)
            {
                ["weekend"] = false,
            },
        };
    }
}
