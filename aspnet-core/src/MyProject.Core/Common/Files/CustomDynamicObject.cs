
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExcelHorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment;
using ExcelVerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment;

namespace MyProject.Common.Files;
using System.Collections.Generic;
using System.Dynamic;

/// <summary>
///     Custom dynamic object class
/// </summary>
public class CustomDynamicObject : DynamicObject
{
    /// <summary>
    ///     The dictionary property used store the data
    /// </summary>
    public Dictionary<string, object> Properties = new();

    /// <summary>
    ///     Provides the implementation for operations that get member values.
    /// </summary>
    /// <param name="binder">Get Member Binder object</param>
    /// <param name="result">The result of the get operation.</param>
    /// <returns>true if the operation is successful; otherwise, false.</returns>
    public override bool TryGetMember(GetMemberBinder binder, out object result)
    {
        result = default;

        if (Properties.ContainsKey(binder.Name))
        {
            result = Properties[binder.Name];
            return true;
        }

        return false;
    }

    /// <summary>
    ///     Provides the implementation for operations that set member values.
    /// </summary>
    /// <param name="binder">Set memeber binder object</param>
    /// <param name="value">The value to set to the member</param>
    /// <returns>true if the operation is successful; otherwise, false.</returns>
    public override bool TrySetMember(SetMemberBinder binder, object value)
    {
        Properties[binder.Name] = value is string
            ? value.NullToEmpty()
                .Replace("\r\n", string.Empty)
                .Replace("\r", string.Empty)
                .Replace("\n", string.Empty)
            : value;
        return true;
    }

    /// <summary>
    ///     Return all dynamic member names
    /// </summary>
    /// <returns>the property name list</returns>
    public override IEnumerable<string> GetDynamicMemberNames()
    {
        return Properties.Keys;
    }

    /// <summary>
    ///     Return data string by key
    /// </summary>
    /// <returns>data</returns>
    public string GetStringValue(string key)
    {
        if (Properties.ContainsKey(key))
        {
            var result = Properties[key].NullToEmpty();

            return result;
        }

        return string.Empty;
    }
}