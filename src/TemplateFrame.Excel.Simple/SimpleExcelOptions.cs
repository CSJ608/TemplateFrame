namespace TemplateFrame.Excel.Simple;

/// <summary>Options for simple-table writing.</summary>
public sealed record SimpleExcelOptions
{
    /// <summary>Worksheet name (default Sheet1).</summary>
    public string SheetName { get; init; } = "Sheet1";

    /// <summary>Whether the header row is bold (default true).</summary>
    public bool BoldHeader { get; init; } = true;

    /// <summary>The starting cell (e.g. "A1" / "C5"); headers land here, data rows follow below.</summary>
    public string StartCell { get; init; } = "A1";

    /// <summary>The named range marking the table area (default TF_Table); empty = no named range.</summary>
    public string TableName { get; init; } = "TF_Table";
}
