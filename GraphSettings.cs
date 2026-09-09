using System.IO;
using System.Text.Json;

namespace ClaimsCalculatorApp;

public sealed class GraphSettings
{
    public string TenantId { get; init; } = "common";
    public string ClientId { get; init; } = "";
    public string ShareUrl { get; init; } = "";
    public string TableName { get; init; } = "ClaimsData";
    public string LocalWorkbookPath { get; init; } = "";

    public static GraphSettings Load()
    {
        const string fileName = "appsettings.json";
        if (!File.Exists(fileName))
            throw new InvalidOperationException($"Missing {fileName}.");

        using var document = JsonDocument.Parse(File.ReadAllText(fileName));
        var graph = document.RootElement.GetProperty("Graph");
        return new GraphSettings
        {
            TenantId = graph.GetProperty("TenantId").GetString() ?? "common",
            ClientId = graph.GetProperty("ClientId").GetString() ?? "",
            ShareUrl = graph.GetProperty("ShareUrl").GetString() ?? "",
            TableName = graph.GetProperty("TableName").GetString() ?? "ClaimsData",
            LocalWorkbookPath = graph.GetProperty("LocalWorkbookPath").GetString() ?? ""
        };
    }
}
