using System;

namespace TemplateFrame.Engine;

/// <summary>
/// 填充选项的通用形状（按缺失必填元素策略枚举类型参数化，供插件类型继承以保留各自公开枚举 API）。
/// <para>English: Common fill-options shape, parameterized by the missing-element policy enum so plugins keep their public enum types.</para>
/// 如 Word 插件：<c>WordFillOptions : TemplateFillOptions&lt;MissingElementPolicy&gt;</c>（迭代 15 公共代码下沉）。
/// </summary>
public abstract record TemplateFillOptions<TMissingPolicy>
    where TMissingPolicy : struct, Enum
{
    /// <summary>缺失必填元素时的处理策略。</summary>
    public TMissingPolicy MissingElementPolicy { get; init; } = default;
}