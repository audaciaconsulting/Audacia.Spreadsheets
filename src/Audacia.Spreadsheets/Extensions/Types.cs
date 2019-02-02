using System;
using System.Linq;
using Audacia.Core.Extensions;

namespace Audacia.Spreadsheets.Extensions
{
    public static class Types
    {
        private static readonly TypeCode[] NumberTypeCodes =
        {
            TypeCode.Byte,
            TypeCode.SByte,
            TypeCode.UInt16,
            TypeCode.UInt32,
            TypeCode.UInt64,
            TypeCode.Int16,
            TypeCode.Int32,
            TypeCode.Int64,
            TypeCode.Decimal,
            TypeCode.Double,
            TypeCode.Single
        };

        /// <summary>
        /// Determines whether a type is numeric
        /// </summary>
        /// <param name="t">Property Type</param> 
        /// <returns>true if numeric</returns>
        public static bool IsNumeric(this Type t)
        {
            var underlyingType = t.GetUnderlyingTypeIfNullable();
            return NumberTypeCodes.Contains(Type.GetTypeCode(underlyingType));
        }
    }
}