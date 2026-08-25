// 低版本目标（netstandard2.0 / net462）的编译垫片：IsExternalInit 支撑 record/init，NotNullWhen 支撑空值流分析注解。
#if !NET5_0_OR_GREATER
namespace System.Runtime.CompilerServices
{
    internal sealed class IsExternalInit { }
}
#endif

#if !NET5_0_OR_GREATER
namespace System.Diagnostics.CodeAnalysis
{
    [AttributeUsage(AttributeTargets.Parameter)]
    internal sealed class NotNullWhenAttribute : Attribute
    {
        public NotNullWhenAttribute(bool returnValue) => ReturnValue = returnValue;

        public bool ReturnValue { get; }
    }
}
#endif
