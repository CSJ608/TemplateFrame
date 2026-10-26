using TemplateFrame.Validation;

namespace TemplateFrame.Services;

internal sealed class MappingWarnings
{
    private readonly List<TemplateValidationIssue> _issues;
    private readonly Dictionary<(string? Key, string? TableKey, int? Row), List<int>> _conversionIndices = new();

    public MappingWarnings(IEnumerable<TemplateValidationIssue> issues)
    {
        _issues = issues.ToList();
        for (var index = 0; index < _issues.Count; index++)
        {
            IndexConversion(index);
        }
    }

    public IReadOnlyList<TemplateValidationIssue> Issues => _issues;

    public void Merge(TemplateValidationIssue warning)
    {
        var index = FindConversion(warning);
        if (index < 0)
        {
            _issues.Add(warning);
            IndexConversion(_issues.Count - 1);
            return;
        }

        _issues[index] = _issues[index] with
        {
            DataPath = warning.DataPath,
            TargetType = warning.TargetType,
        };
    }

    private int FindConversion(TemplateValidationIssue warning)
    {
        if (_conversionIndices.TryGetValue(Location(warning), out var indices))
        {
            foreach (var index in indices)
            {
                var existing = _issues[index];
                if (IsConversion(existing) && Equals(existing.RawValue, warning.RawValue))
                {
                    return index;
                }
            }
        }

        return -1;
    }

    private void IndexConversion(int index)
    {
        var issue = _issues[index];
        if (!IsConversion(issue))
        {
            return;
        }

        var location = Location(issue);
        if (!_conversionIndices.TryGetValue(location, out var indices))
        {
            indices = [];
            _conversionIndices.Add(location, indices);
        }

        indices.Add(index);
    }

    private static bool IsConversion(TemplateValidationIssue issue)
        => issue.Code == TemplateValidationIssueCode.ConversionFailed && !string.IsNullOrEmpty(issue.TargetType);

    private static (string? Key, string? TableKey, int? Row) Location(TemplateValidationIssue issue)
        => (issue.Key, issue.TableKey, issue.DataRowNumber);
}
