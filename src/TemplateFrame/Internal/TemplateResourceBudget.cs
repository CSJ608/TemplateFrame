using TemplateFrame.Engine;
using TemplateFrame.Localization;

namespace TemplateFrame.Internal;

/// <summary>Cumulative table work budget owned by one operation, never shared between calls.</summary>
internal sealed class TemplateResourceBudget
{
    private readonly TemplateResourceLimits _limits;
    private long _rows;
    private long _cells;
    private long _clonedXmlCharacters;

    public TemplateResourceBudget(TemplateResourceLimits limits)
    {
        Guard.ThrowIfNull(limits, nameof(limits));
        limits.Validate();
        _limits = limits;
    }

    public void AddClonedXml(long templateCharacters, long copies)
    {
        if (templateCharacters < 0 || copies < 0)
            throw new ArgumentOutOfRangeException(templateCharacters < 0 ? nameof(templateCharacters) : nameof(copies));
        if (templateCharacters > 0 && copies > (_limits.MaxClonedXmlCharacters - _clonedXmlCharacters) / templateCharacters)
            throw new InvalidOperationException(Sr.Get("ResourceLimit.Exceeded", nameof(_limits.MaxClonedXmlCharacters)));
        _clonedXmlCharacters += templateCharacters * copies;
    }

    public void AddCells(long cells)
    {
        if (cells < 0)
            throw new ArgumentOutOfRangeException(nameof(cells));
        if (cells > _limits.MaxTableCells - _cells)
            throw new InvalidOperationException(Sr.Get("ResourceLimit.Exceeded", nameof(_limits.MaxTableCells)));
        _cells += cells;
    }

    public void AddRows(long rows, long columns)
    {
        if (rows < 0 || columns < 0)
            throw new ArgumentOutOfRangeException(rows < 0 ? nameof(rows) : nameof(columns));
        if (rows > _limits.MaxTableRows - _rows)
            throw new InvalidOperationException(Sr.Get("ResourceLimit.Exceeded", nameof(_limits.MaxTableRows)));
        if (columns > 0 && rows > (_limits.MaxTableCells - _cells) / columns)
            throw new InvalidOperationException(Sr.Get("ResourceLimit.Exceeded", nameof(_limits.MaxTableCells)));
        _rows += rows;
        _cells += rows * columns;
    }
}
