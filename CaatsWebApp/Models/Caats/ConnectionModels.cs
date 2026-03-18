namespace CaatsWebApp.Models.Caats;

public sealed class ConnectRequest
{
    public string Server { get; set; } = string.Empty;
    public string Driver { get; set; } = "ODBC Driver 17 for SQL Server";
    public bool TrustedConnection { get; set; } = true;
    public string Username { get; set; } = string.Empty;
    public string Password { get; set; } = string.Empty;
}

public sealed class ConnectResponse
{
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
    public List<string> Databases { get; set; } = [];
}

public sealed class UseDatabaseRequest
{
    public string Database { get; set; } = string.Empty;
}

public sealed class GenericResponse
{
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
}
