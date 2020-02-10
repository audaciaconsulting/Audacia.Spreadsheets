using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using Audacia.Core.Extensions;
using Audacia.Spreadsheets.Attributes;

namespace Audacia.Spreadsheets.Extensions
{
    internal static class Types
    {
        /// <summary>
        /// Returns an enumerable of properties that should display on the worksheet.
        /// </summary>
        public static IEnumerable<PropertyInfo> GetProps(this Type classType, params string[] ignoreProperties)
        {
            return classType
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(property =>
                {
                    // DataMember should be ignored by default everywhere
                    var ignoreDataMemberAttrs = property.GetCustomAttributes(typeof(IgnoreDataMemberAttribute), false);
                    if (ignoreDataMemberAttrs.Any())
                    {
                        return false;
                    }

                    // Property should ony be ignored by the spreadsheet exporter
                    var cellIgnoreAttrs = property.GetCustomAttributes(typeof(CellIgnoreAttribute), false);
                    if (cellIgnoreAttrs.Any())
                    {
                        return false;
                    }

                    // Property is a primary key and should be ignored (Legacy support)
                    var primaryKeyAttrs = property.GetCustomAttributes(typeof(IdColumnAttribute), false);
                    if (primaryKeyAttrs.Any())
                    {
                        return false;
                    }

                    // Property should be ignored if it's in the list of ignoreProperties
                    return !ignoreProperties?.Contains(property.Name) ?? true;
                })
                .Where(p =>
                {
                    // And the property is a value type not a class
                    var underlyingType = p.PropertyType.GetUnderlyingTypeIfNullable();
                    return underlyingType.IsValueType || underlyingType == typeof(string);
                });
        }

        /* Keeping on the off chance we still want to use this when creating TableColumns...
           We were originally using this, and then this calls the method above. 
        internal static IEnumerable<PropertyInfo> GetBaseProps(this Type classType, params string[] ignoreProperties)
        {
            var baseType = classType.BaseType;
            if (baseType != null && baseType.IsClass)
            {
                var parentProperties = GetBaseProps(baseType, ignoreProperties).ToArray();
                var parentPropertyNames = parentProperties.Select(p => p.Name).ToArray();
                var childProperties = GetProps(classType, ignoreProperties)
                    .Where(p => !parentPropertyNames.Contains(p.Name)).ToArray();

                return parentProperties.Union(childProperties);
            }

            return GetProps(classType, ignoreProperties);
        }
        */
    }
}