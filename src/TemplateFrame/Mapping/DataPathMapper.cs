using System.Collections;
using System.Collections.Concurrent;
using System.Runtime.CompilerServices;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Localization;

namespace TemplateFrame.Mapping;

public static class DataPathMapper
{
    private static readonly ConditionalWeakTable<TemplateContract, ConcurrentDictionary<Type, ContractMapping>> Cache = new();

    public static FillData ToFillData<TData>(TData data, TemplateContract contract)
    {
        Guard.ThrowIfNull(data, nameof(data));
        Guard.ThrowIfNull(contract, nameof(contract));
        var mapping = GetMapping(contract, typeof(TData));
        var values = new Dictionary<string, object?>();
        var tables = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>();

        foreach (var element in contract.Elements)
        {
            switch (element)
            {
                case TemplateElement scalar when (scalar is TextElement or ImageElement) && scalar.DataPath is { Length: > 0 }:
                    values[scalar.Key] = mapping.Scalars[scalar.Key].GetValue(data);
                    break;

                case TableElement table when table.DataPath is { Length: > 0 } || mapping.Tables.ContainsKey(table.Key):
                    tables[table.Key] = MapTableToData(mapping.Tables[table.Key], data);
                    break;
            }
        }

        return new FillData { Values = values, Tables = tables };
    }

    public static TData FromFillData<TData>(FillData data, TemplateContract contract)
        => FromFillData<TData>(data, contract, lenientConversion: false);

    internal static TData FromFillData<TData>(FillData data, TemplateContract contract, bool lenientConversion)
    {
        Guard.ThrowIfNull(data, nameof(data));
        Guard.ThrowIfNull(contract, nameof(contract));
        var mapping = GetMapping(contract, typeof(TData));
        var assignment = new PropertyAssignment(lenientConversion, data.MappingWarning);
        object? instance = null;

        foreach (var element in contract.Elements)
        {
            switch (element)
            {
                case TemplateElement scalar when (scalar is TextElement or ImageElement) && scalar.DataPath is { Length: > 0 }
                    && data.Values.TryGetValue(scalar.Key, out var scalarValue):
                    instance ??= CreateInstance(typeof(TData));
                    assignment.Set(instance, scalarValue,
                        new MappingTarget(mapping.Scalars[scalar.Key], scalar.Key, scalar.DataPath!, Format: (scalar as TextElement)?.Format));
                    break;

                case TableElement table when mapping.Tables.TryGetValue(table.Key, out var tableMapping):
                    {
                        var collectionProperty = tableMapping.CollectionProperty;
                        if (collectionProperty is null)
                        {
                            data.Tables.TryGetValue(table.Key, out var rootRows);
                            return (TData)MapTableFromData(tableMapping, rootRows ?? [], typeof(TData), assignment, table);
                        }

                        if (data.Tables.TryGetValue(table.Key, out var rows))
                        {
                            instance ??= CreateInstance(typeof(TData));
                            collectionProperty.SetValue(
                                instance,
                                MapTableFromData(tableMapping, rows, collectionProperty.PropertyType, assignment, table));
                        }

                        break;
                    }
            }
        }

        return (TData)(instance ?? CreateInstance(typeof(TData)));
    }

    private static object CreateInstance(Type type)
        => Activator.CreateInstance(type)
           ?? throw new InvalidOperationException(Sr.Get("Mapping.CreateInstanceFailed", type.Name));

    private static ContractMapping GetMapping(TemplateContract contract, Type dataType)
    {
        var mappings = Cache.GetValue(contract, static _ => new ConcurrentDictionary<Type, ContractMapping>());
        return mappings.GetOrAdd(dataType, type => ContractMapping.Create(contract, type));
    }

    public static bool IsCollectionDataType(Type type)
        => ContractMapping.IsCollectionDataType(type);

    private static IReadOnlyList<IReadOnlyDictionary<string, object?>> MapTableToData(
        TableMapping tableMapping,
        object? data)
    {
        var result = new List<IReadOnlyDictionary<string, object?>>();
        var lines = tableMapping.CollectionProperty is null
            ? data as IEnumerable
            : tableMapping.CollectionProperty.GetValue(data) as IEnumerable;
        if (lines is null)
        {
            return result;
        }

        foreach (var line in lines)
        {
            var row = new Dictionary<string, object?>();
            foreach (var column in tableMapping.Columns)
            {
                row[column.Key] = column.Value.GetValue(line);
            }

            result.Add(row);
        }

        return result;
    }

    private static object MapTableFromData(
        TableMapping tableMapping,
        IReadOnlyList<IReadOnlyDictionary<string, object?>> rows,
        Type destinationType,
        PropertyAssignment assignment,
        TableElement table)
    {
        var elementType = tableMapping.ElementType;
        if (destinationType.IsArray)
        {
            var array = Array.CreateInstance(elementType, rows.Count);
            for (var i = 0; i < rows.Count; i++)
            {
                array.SetValue(CreateLine(tableMapping, rows[i], assignment, table, i), i);
            }

            return array;
        }

        var list = (IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(elementType))!;
        for (var i = 0; i < rows.Count; i++)
        {
            list.Add(CreateLine(tableMapping, rows[i], assignment, table, i));
        }

        return list;
    }

    private static object CreateLine(
        TableMapping tableMapping,
        IReadOnlyDictionary<string, object?> row,
        PropertyAssignment assignment,
        TableElement table,
        int rowIndex)
    {
        var line = Activator.CreateInstance(tableMapping.ElementType)!;
        foreach (var column in tableMapping.Columns)
        {
            if (row.TryGetValue(column.Key, out var value))
            {
                assignment.Set(line, value,
                    new MappingTarget(column.Value, column.Key, table.DataPath ?? string.Empty, table.Key, rowIndex + 1));
            }
        }

        return line;
    }

}
