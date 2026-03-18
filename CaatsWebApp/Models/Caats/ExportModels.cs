namespace CaatsWebApp.Models.Caats;

public sealed class ExportRequest
{
    public string OutputFolder { get; set; } = string.Empty;
    public bool Word { get; set; } = true;
    public bool Excel { get; set; } = true;
    public bool Csv { get; set; }
}

public sealed class ExportResponse
{
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
    public List<string> Files { get; set; } = [];
}
