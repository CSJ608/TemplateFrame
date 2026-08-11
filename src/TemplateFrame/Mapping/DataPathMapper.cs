using System.Collections;
using System.Collections.Concurrent;
using System.Globalization;
using System.Reflection;
using TemplateFrame.Contract;
using TemplateFrame.Data;
using TemplateFrame.Localization;

namespace TemplateFrame.Mapping;

/// <summary>
/// 自动映射器（迭代 9）：按契约元素的 <see cref="TemplateElement.DataPath"/> 反射完成 TData ⇄ FillData 双向映射。
/// <para>English: Auto-mapper — reflects TData ⇄ FillData by the DataPath declared on contract elements.</para>
/// 显式 DataPath 为主：标量 / 图片用单级属性路径，表格用「集合属性 + 列属性」两级路径；
/// 数据对象本身就是集合（<c>List&lt;T&gt;</c> / <c>IReadOnlyList&lt;T&gt;</c> / 数组）时，表格 DataPath 留空即按「根集合」映射。
/// 未声明 DataPath 的元素不参与自动映射（可对个别字段手写映射，或重写业务服务的映射方法）。
/// 属性解析按（契约, 数据类型）缓存，只在首次映射时反射一次。
/// </summary>
public static class DataPathMapper
{
    private static readonly ConcurrentDictionary<(TemplateContract Contract, Type DataType), ContractMapping> Cache = new();

    /// <summary>正向：强类型数据 → <see cref="FillData"/>（只映射声明了 DataPath 的元素）。</summary>
    public static FillData ToFillData<TData>(TData data, TemplateContract contract)
    {
        ArgumentNullException.ThrowIfNull(contract);
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

    /// <summary>
    /// 反向：<see cref="FillData"/> → 强类型数据（容器类型需有无参构造；缺失/空值字段保持默认）。
    /// 根集合模式（TData 本身是集合、表格 DataPath 留空）直接返回行集合；接口集合
    /// （<c>IReadOnlyList&lt;T&gt;</c> / <c>IEnumerable&lt;T&gt;</c>）由 <c>List&lt;T&gt;</c> 承载。
    /// </summary>
    public static TData FromFillData<TData>(FillData data, TemplateContract contract)
    {
        ArgumentNullException.ThrowIfNull(data);
        ArgumentNullException.ThrowIfNull(contract);
        var mapping = GetMapping(contract, typeof(TData));
        object? instance = null;

        foreach (var element in contract.Elements)
        {
            switch (element)
            {
                case TemplateElement scalar when (scalar is TextElement or ImageElement) && scalar.DataPath is { Length: > 0 }
                    && data.Values.TryGetValue(scalar.Key, out var scalarValue):
                    instance ??= CreateInstance(typeof(TData));
                    SetValue(mapping.Scalars[scalar.Key], instance, scalarValue, scalar as TextElement);
                    break;

                case TableElement table when mapping.Tables.TryGetValue(table.Key, out var tableMapping):
                {
                    var collectionProperty = tableMapping.CollectionProperty;
                    if (collectionProperty is null)
                    {
                        // 根集合：直接返回行集合（缺数据时返回空集合）
                        data.Tables.TryGetValue(table.Key, out var rootRows);
                        return (TData)MapTableFromData(tableMapping, rootRows ?? [], typeof(TData));
                    }

                    if (data.Tables.TryGetValue(table.Key, out var rows))
                    {
                        instance ??= CreateInstance(typeof(TData));
                        collectionProperty.SetValue(
                            instance,
                            MapTableFromData(tableMapping, rows, collectionProperty.PropertyType));
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
        => Cache.GetOrAdd((contract, dataType), static key => BuildMapping(key.Contract, key.DataType));

    private static ContractMapping BuildMapping(TemplateContract contract, Type dataType)
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
                    // 根集合模式：TData 本身就是集合（List<T> / IReadOnlyList<T> / 数组），表格 DataPath 留空，行数据直接取根对象
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

    /// <summary>把表格列声明（列 DataPath = 行元素属性）解析进 <paramref name="tableMapping"/>。</summary>
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

    /// <summary>
    /// 类型是否为可映射的集合（数组或泛型 <c>IEnumerable&lt;T&gt;</c>；string 除外）。
    /// 根集合模式（数据对象本身就是集合）据此识别。
    /// </summary>
    public static bool IsCollectionDataType(Type type)
    {
        ArgumentNullException.ThrowIfNull(type);
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
            foreach (var (columnKey, property) in tableMapping.Columns)
            {
                row[columnKey] = property.GetValue(line);
            }

            result.Add(row);
        }

        return result;
    }

    /// <summary>
    /// 行集合 → 目标集合类型（根集合或集合属性）。数组返回 <c>T[]</c>，其余返回 <c>List&lt;T&gt;</c>
    /// （可赋值给 <c>IReadOnlyList&lt;T&gt;</c> / <c>IEnumerable&lt;T&gt;</c> 等接口）。
    /// </summary>
    private static object MapTableFromData(
        TableMapping tableMapping,
        IReadOnlyList<IReadOnlyDictionary<string, object?>> rows,
        Type destinationType)
    {
        var elementType = tableMapping.ElementType;
        if (destinationType.IsArray)
        {
            var array = Array.CreateInstance(elementType, rows.Count);
            for (var i = 0; i < rows.Count; i++)
            {
                array.SetValue(CreateLine(tableMapping, rows[i]), i);
            }

            return array;
        }

        var list = (IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(elementType))!;
        foreach (var row in rows)
        {
            list.Add(CreateLine(tableMapping, row));
        }

        return list;
    }

    private static object CreateLine(
        TableMapping tableMapping,
        IReadOnlyDictionary<string, object?> row)
    {
        var line = Activator.CreateInstance(tableMapping.ElementType)!;
        foreach (var (columnKey, property) in tableMapping.Columns)
        {
            if (row.TryGetValue(columnKey, out var value))
            {
                SetValue(property, line, value, null);
            }
        }

        return line;
    }

    private static void SetValue(PropertyInfo property, object instance, object? value, TextElement? element)
    {
        var targetType = property.PropertyType;
        var converted = ConvertValue(value, targetType, element);
        if (converted is null && targetType.IsValueType && Nullable.GetUnderlyingType(targetType) is null)
        {
            // 空值落在非空值类型：保持构造默认值（跳过赋值）
            return;
        }

        property.SetValue(instance, converted);
    }

    /// <summary>把 FillData 值转换为目标属性类型（含数值收敛：SimpleExcel 读回的数字是 double）。</summary>
    private static object? ConvertValue(object? value, Type targetType, TextElement? element)
    {
        if (value is null)
        {
            return !targetType.IsValueType || Nullable.GetUnderlyingType(targetType) is not null
                ? null
                : Activator.CreateInstance(targetType);
        }

        // 空字符串（如 Word 回读的空日期/数字单元格）视为空值；string 目标保持原样
        if (value is string emptyText && string.IsNullOrWhiteSpace(emptyText) && targetType != typeof(string))
        {
            return null;
        }

        if (targetType.IsInstanceOfType(value))
        {
            return value;
        }

        if (targetType == typeof(string))
        {
            return value switch
            {
                DateTime dateTime => dateTime.ToString(element?.Format, CultureInfo.InvariantCulture),
                IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture),
                _ => value.ToString(),
            };
        }

        var underlying = Nullable.GetUnderlyingType(targetType) ?? targetType;
        if (underlying == typeof(decimal))
        {
            return value is string s
                ? decimal.Parse(s, NumberStyles.Number, CultureInfo.InvariantCulture)
                : Convert.ToDecimal(value, CultureInfo.InvariantCulture);
        }

        if (underlying == typeof(int))
        {
            return Convert.ToInt32(value, CultureInfo.InvariantCulture);
        }

        if (underlying == typeof(long))
        {
            return Convert.ToInt64(value, CultureInfo.InvariantCulture);
        }

        if (underlying == typeof(float))
        {
            return Convert.ToSingle(value, CultureInfo.InvariantCulture);
        }

        if (underlying == typeof(double))
        {
            return Convert.ToDouble(value, CultureInfo.InvariantCulture);
        }

        if (underlying == typeof(bool))
        {
            return value is string boolText
                ? bool.Parse(boolText)
                : Convert.ToBoolean(value, CultureInfo.InvariantCulture);
        }

        if (underlying == typeof(DateTime))
        {
            if (value is string dateText)
            {
                return element?.Format is { Length: > 0 }
                    ? DateTime.ParseExact(dateText, element.Format, CultureInfo.InvariantCulture)
                    : DateTime.Parse(dateText, CultureInfo.InvariantCulture);
            }

            return Convert.ToDateTime(value, CultureInfo.InvariantCulture);
        }

        if (underlying == typeof(byte[]))
        {
            return value is byte[] bytes
                ? bytes
                : Convert.FromBase64String(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
        }

        throw new InvalidOperationException(
            Sr.Get("Mapping.ConvertFailed", value.GetType().Name, targetType.Name));
    }

    /// <summary>单个契约（+ 数据类型）解析一次得到的映射模型。</summary>
    private sealed class ContractMapping
    {
        /// <summary>元素 Key → 属性（文本/图片）。</summary>
        public Dictionary<string, PropertyInfo> Scalars { get; } = new(StringComparer.Ordinal);

        /// <summary>表格 Key → 表格映射（集合属性 + 列属性）。</summary>
        public Dictionary<string, TableMapping> Tables { get; } = new(StringComparer.Ordinal);
    }

    /// <summary>表格映射：集合属性（根集合模式为 null）+ 行元素类型 + 列 Key → 行属性。</summary>
    private sealed class TableMapping
    {
        /// <summary>集合属性；为 null 表示根集合模式（数据对象本身就是集合）。</summary>
        public PropertyInfo? CollectionProperty { get; init; }

        public Type ElementType { get; init; } = null!;

        public Dictionary<string, PropertyInfo> Columns { get; } = new(StringComparer.Ordinal);
    }
}
