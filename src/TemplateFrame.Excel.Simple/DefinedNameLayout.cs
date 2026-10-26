namespace TemplateFrame.Excel.Simple;

internal sealed record DefinedNameLayout(
    SimpleExcelAddress.TableRange Range,
    IReadOnlyList<string> Headers,
    IReadOnlyDictionary<string, int> ColumnIndex,
    IReadOnlyList<string> AmbiguousColumnKeys,
    IReadOnlyList<string> PositionAmbiguousColumnKeys);
