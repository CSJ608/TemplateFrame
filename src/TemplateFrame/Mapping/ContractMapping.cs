using System.Collections;
using System.Reflection;
using TemplateFrame.Contract;
using TemplateFrame.Localization;

namespace TemplateFrame.Mapping;

internal sealed class ContractMapping
{
    public Dictionary<string, PropertyInfo> Scalars { get; } = new(StringComparer.Ordinal);
    public Dictionary<string, TableMapping> Tables { get; } = new(StringComparer.Ordinal);

    public static ContractMapping Create(TemplateContract contract, Type dataType)
    {
        var mapping = new ContractMapping();
        var usedScalarProperties = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var usedTableProperties = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var isCollectionDataType = IsCollectionDataType(dataType);
        var rootTableSeen = false;

        foreach (var element in contract.Elements)
        {
            switch (element)
            {
                case TemplateElement scalar when (scalar is TextElement or ImageElement) && scalar.DataPath is { Length: > 0 }:
                    {
                        var property = FindProperty(dataType, scalar.DataPath!, contract, scalar);
                        if (!usedScalarProperties.Add(property.Name))
                        {
                            throw new InvalidOperationException(
                                Sr.Get("Mapping.DuplicateScalarDataPath", contract.Name, property.Name));
                        }

                        mapping.Scalars[scalar.Key] = property;
                        break;
                    }

                case TableElement table when table.DataPath is { Length: > 0 } || isCollectionDataType:
                    {
                        var isRootCollection = string.IsNullOrWhiteSpace(table.DataPath);
                        if (isRootCollection)
                        {
                            if (rootTableSeen)
                            {
                                throw new InvalidOperationException(
                                    Sr.Get("Mapping.RootCollectionMultipleTables", contract.Name));
                            }

                            rootTableSeen = true;
                            var rootElementType = GetCollectionElementType(dataType, contract, table);
                            var rootMapping = new TableMapping { ElementType = rootElementType };
                            BuildColumns(rootMapping, rootElementType, contract, table);
                            mapping.Tables[table.Key] = rootMapping;
                            break;
                        }

                        if (isCollectionDataType)
                        {
                            throw new InvalidOperationException(
                                Sr.Get("Mapping.RootCollectionHasDataPath", contract.Name, table.Key, table.DataPath));
                        }

                        var property = FindProperty(dataType, table.DataPath!, contract, table);
                        if (!usedTableProperties.Add(property.Name))
                        {
                            throw new InvalidOperationException(
                                Sr.Get("Mapping.DuplicateTableDataPath", contract.Name, property.Name));
                        }

                        var elementType = GetCollectionElementType(property.PropertyType, contract, table);
                        var tableMapping = new TableMapping { CollectionProperty = property, ElementType = elementType };
                        BuildColumns(tableMapping, elementType, contract, table);
                        mapping.Tables[table.Key] = tableMapping;
                        break;
                    }
            }
        }

        return mapping;
    }

    private static void BuildColumns(
        TableMapping tableMapping,
        Type elementType,
        TemplateContract contract,
        TableElement table)
    {
        var usedColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var column in table.Columns)
        {
            if (column.DataPath is not { Length: > 0 })
            {
                continue;
            }

            var columnProperty = FindProperty(elementType, column.DataPath!, contract, column);
            if (!usedColumns.Add(columnProperty.Name))
            {
                throw new InvalidOperationException(
                    Sr.Get("Mapping.DuplicateColumnDataPath", contract.Name, table.Key, columnProperty.Name));
            }

            tableMapping.Columns[column.Key] = columnProperty;
        }
    }

    private static PropertyInfo FindProperty(Type type, string path, TemplateContract contract, TemplateElement element)
    {
        var property = type.GetProperty(path, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
        if (property is null)
        {
            throw new InvalidOperationException(
                Sr.Get("Mapping.PropertyNotFound", contract.Name, element.Key, element.DisplayName, path, type.Name));
        }

        if (!property.CanWrite)
        {
            throw new InvalidOperationException(
                Sr.Get("Mapping.ReadOnlyProperty", contract.Name, element.Key, path, type.Name, property.Name));
        }

        return property;
    }

    public static bool IsCollectionDataType(Type type)
    {
        Guard.ThrowIfNull(type, nameof(type));
        return type != typeof(string)
               && (type.IsArray
                   || type.GetInterfaces()
                       .Concat([type])
                       .Any(t => t.IsGenericType && t.GetGenericTypeDefinition() == typeof(IEnumerable<>)));
    }

    private static Type GetCollectionElementType(Type propertyType, TemplateContract contract, TableElement table)
    {
        if (propertyType == typeof(string))
        {
            throw new InvalidOperationException(
                Sr.Get("Mapping.TablePointsToString", contract.Name, table.Key, table.DataPath));
        }

        if (propertyType.IsArray)
        {
            return propertyType.GetElementType()!;
        }

        var enumerable = propertyType.GetInterfaces()
            .Concat([propertyType])
            .FirstOrDefault(t => t.IsGenericType && t.GetGenericTypeDefinition() == typeof(IEnumerable<>));
        if (enumerable is not null)
        {
            return enumerable.GetGenericArguments()[0];
        }

        if (typeof(IEnumerable).IsAssignableFrom(propertyType))
        {
            throw new InvalidOperationException(
                Sr.Get("Mapping.TablePointsToNonGenericCollection", contract.Name, table.Key, table.DataPath, propertyType.Name));
        }

        throw new InvalidOperationException(
            Sr.Get("Mapping.TablePointsToNonCollection", contract.Name, table.Key, table.DataPath, propertyType.Name));
    }

}

internal sealed class TableMapping
{
    public PropertyInfo? CollectionProperty { get; init; }
    public Type ElementType { get; init; } = null!;
    public Dictionary<string, PropertyInfo> Columns { get; } = new(StringComparer.Ordinal);
}
