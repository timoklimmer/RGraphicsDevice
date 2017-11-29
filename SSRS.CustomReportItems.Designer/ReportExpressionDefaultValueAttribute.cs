using System;
using System.ComponentModel;
using System.Globalization;
using Microsoft.ReportingServices.RdlObjectModel;

namespace SSRS.CustomReportItems.Designer
{
    /// <summary>
    ///     Specifies a default value for a <see cref="ReportExpression" /> property.
    /// </summary>
    public class ReportExpressionDefaultValueAttribute : DefaultValueAttribute
    {
        /// <summary>
        ///     Constructor.
        /// </summary>
        public ReportExpressionDefaultValueAttribute()
            : base(new ReportExpression())
        {
        }

        /// <summary>
        ///     Constructor.
        /// </summary>
        public ReportExpressionDefaultValueAttribute(string value)
            : base(new ReportExpression(value))
        {
        }

        /// <summary>
        ///     Constructor.
        /// </summary>
        public ReportExpressionDefaultValueAttribute(Type type)
            : base(Activator.CreateInstance(ConstructGenericType(type)))
        {
        }

        /// <summary>
        ///     Constructor.
        /// </summary>
        public ReportExpressionDefaultValueAttribute(Type type, object value)
            : base(CreateInstance(type, value))
        {
        }

        /// <summary>
        ///     Constructs a generic <see cref="ReportExpression" /> of the specified type.
        /// </summary>
        internal static Type ConstructGenericType(Type type)
        {
            return typeof (ReportExpression<>).MakeGenericType(new[]
            {
                type
            });
        }

        /// <summary>
        ///     Compiler-optimized method to create a certain instance.
        /// </summary>
        internal static object CreateInstance(Type type, object value)
        {
            type = ConstructGenericType(type);
            if (value is string)
                return type.GetConstructor(new[]
                {
                    typeof (string),
                    typeof (IFormatProvider)
                }).Invoke(new[]
                {
                    value,
                    CultureInfo.InvariantCulture
                });
            return Activator.CreateInstance(type, new[]
            {
                value
            });
        }
    }
}