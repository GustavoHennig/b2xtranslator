#if !NET6_0_OR_GREATER
using System;

namespace System.Diagnostics.CodeAnalysis
{
    [Flags]
    public enum DynamicallyAccessedMemberTypes
    {
        None = 0,
        PublicConstructors = 8
    }

    [AttributeUsage(AttributeTargets.GenericParameter | AttributeTargets.Parameter | AttributeTargets.Field | AttributeTargets.Property | AttributeTargets.ReturnValue, Inherited = false)]
    public sealed class DynamicallyAccessedMembersAttribute : Attribute
    {
        public DynamicallyAccessedMembersAttribute(DynamicallyAccessedMemberTypes memberTypes)
        {
            this.MemberTypes = memberTypes;
        }

        public DynamicallyAccessedMemberTypes MemberTypes { get; }
    }

    [AttributeUsage(AttributeTargets.All, Inherited = false, AllowMultiple = true)]
    public sealed class UnconditionalSuppressMessageAttribute : Attribute
    {
        public UnconditionalSuppressMessageAttribute(string category, string checkId)
        {
            this.Category = category;
            this.CheckId = checkId;
        }

        public string Category { get; }

        public string CheckId { get; }

        public string Justification { get; set; }

        public string MessageId { get; set; }

        public string Scope { get; set; }

        public string Target { get; set; }
    }

    [AttributeUsage(AttributeTargets.Constructor | AttributeTargets.Method | AttributeTargets.Class | AttributeTargets.Struct, Inherited = false)]
    public sealed class RequiresUnreferencedCodeAttribute : Attribute
    {
        public RequiresUnreferencedCodeAttribute(string message)
        {
            this.Message = message;
        }

        public string Message { get; }

        public string Url { get; set; }
    }
}
#endif
