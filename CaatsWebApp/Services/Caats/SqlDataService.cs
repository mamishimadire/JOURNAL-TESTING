using System.Data;
using Microsoft.Data.SqlClient;

namespace CaatsWebApp.Services.Caats;

public sealed class SqlDataService
{
    private const int SqlCommandTimeoutSeconds = 180;

    public async Task<List<string>> GetDatabasesAsync(string connectionString)
    {
        await using var conn = new SqlConnection(connectionString);
        await conn.OpenAsync();
        var cmd = new SqlCommand("SELECT name FROM sys.databases WHERE name NOT IN ('master','tempdb','model','msdb') ORDER BY name", conn)
        {
            CommandTimeout = SqlCommandTimeoutSeconds,
        };
        var dbs = new List<string>();
        await using var rdr = await cmd.ExecuteReaderAsync();
        while (await rdr.ReadAsync())
        {
            dbs.Add(rdr.GetString(0));
        }

        return dbs;
    }

    public async Task<List<string>> GetTablesAsync(string connectionString)
    {
        await using var conn = new SqlConnection(connectionString);
        await conn.OpenAsync();
        var cmd = new SqlCommand("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' ORDER BY TABLE_NAME", conn)
        {
            CommandTimeout = SqlCommandTimeoutSeconds,
        };
        var tables = new List<string>();
        await using var rdr = await cmd.ExecuteReaderAsync();
        while (await rdr.ReadAsync())
        {
            tables.Add(rdr.GetString(0));
        }

        return tables;
    }

    public async Task<DataTable> GetTableAsync(string connectionString, string tableName, int? top = null)
    {
        await using var conn = new SqlConnection(connectionString);
        await conn.OpenAsync();

        var escaped = tableName.Replace("]", "]]", StringComparison.Ordinal);
        var sql = top.HasValue ? $"SELECT TOP ({top.Value}) * FROM [{escaped}]" : $"SELECT * FROM [{escaped}]";
        var cmd = new SqlCommand(sql, conn)
        {
            CommandTimeout = SqlCommandTimeoutSeconds,
        };
        var dt = new DataTable();
        await using var rdr = await cmd.ExecuteReaderAsync();
        dt.Load(rdr);
        return dt;
    }

    public async Task<List<Dictionary<string, object?>>> ToDictionaryRowsAsync(string connectionString, string tableName, int top)
    {
        var dt = await GetTableAsync(connectionString, tableName, top);
        return dt.AsEnumerable()
            .Select(r => dt.Columns.Cast<DataColumn>().ToDictionary(c => c.ColumnName, c => r[c] == DBNull.Value ? null : r[c]))
            .ToList();
    }
}
