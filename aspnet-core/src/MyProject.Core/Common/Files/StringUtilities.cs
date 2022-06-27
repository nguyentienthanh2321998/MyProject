
using Microsoft.AspNetCore.Http;
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
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using ExcelHorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment;
using ExcelVerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment;

namespace MyProject.Common.Files;

public static class StringUtilities
{
    private static readonly Random _random = new();

    private static readonly int[] _UnicodeCharactersList =
        Enumerable.Range(48, 10) // Numbers           48   - 57
            .Concat(Enumerable.Range(65, 26)) // English uppercase 65   - 90
            .ToArray();

    public static string GenerateOTP()
    {
        var token = new Random().Next(0, 999999).ToString("000000");

        return token;
    }

    public static string GetKeywordElasticFullTextSearch(this string value)
    {
        var result = string.Join(
            " ",
            value.Split(' ').Select(item => string.Format("*{0}*", item)));

        return result;
    }

    public static string GetKeywordElasticFullTextSearchV2(this string value)
    {
        var arrWildcard = value.Split(' ');
        var temp = "";
        for (var i = 0; i < arrWildcard.Length; i++)
            if (i < arrWildcard.Length - 1)
                temp = temp + arrWildcard[i] + " AND ";
            else
                temp = temp + arrWildcard[i];

        var result = "(*" + value + ") OR " +
                     "(" + value + "*) OR " +
                     "(" + temp + ") OR " +
                     "(" + string.Join(" ", value.Split(' ').Select(item => string.Format("{0}*", item))) + ")";

        return result;
    }

    public static string GetValueWithoutSpecialCharacters(
        this string value,
        string separatorCharacter = "")
    {
        var patternRegex = @"[^a-zA-Z0-9_]+";
        var result = Regex.Replace(value.UnicodeToASCII(), patternRegex, separatorCharacter);

        return result.ToLower();
    }

    public static string NullToEmpty(this object value)
    {
        if (value == null || value == DBNull.Value) return string.Empty;

        return value.ToString().Trim();
    }

    public static string RandomAppId(int length = 8)
    {
        Random random = new Random();
        const string chars = "ABCDEF0123456789";
        return new string(Enumerable.Repeat(chars, length)
            .Select(s => s[random.Next(s.Length)]).ToArray());
    }

    public static string GenarateRandomString(int maxSize = 12)
    {
        // Step 1: Random number
        var randomNumberWithDate = new Random((int)DateTime.Now.Ticks)
            .Next(1, 1000)
            .ToString();

        // Step 2: Random difference sequence
        var differenceString = string.Empty;

        var difference = maxSize - randomNumberWithDate.Length;
        for (var i = 0; i < difference; i++)
        {
            var randomCharacter = _UnicodeCharactersList[
                _random.Next(1, _UnicodeCharactersList.Length)];

            differenceString += Convert.ToChar(randomCharacter);
        }

        // Step 3: Insert the string in Step 1 anywhere from the string in Step 2
        var index = new Random().Next(0, differenceString.Length);
        var result = differenceString
            .Insert(index, randomNumberWithDate);

        return result;
    }

    public static string UnicodeToASCII(this string value)
    {
        if (string.IsNullOrEmpty(value))
            return string.Empty;

        var regex = new Regex("\\p{IsCombiningDiacriticalMarks}+");
        var temp = value.Normalize(NormalizationForm.FormD);

        return regex.Replace(temp, string.Empty).Replace('đ', 'd').Replace('Đ', 'D');
    }

    public static string CreateMD5(this Stream stream)
    {
        // Use input string to calculate MD5 hash
        using (var create = MD5.Create())
        {
            var hashBytes = create.ComputeHash(stream);

            // Convert the byte array to hexadecimal string
            var builder = new StringBuilder();
            for (var i = 0; i < hashBytes.Length; i++) builder.Append(hashBytes[i].ToString("X2").ToLower());

            return builder.ToString();
        }
    }

    public static string Truncate(this string value, int maxLength)
    {
        if (!string.IsNullOrEmpty(value))
            return value.Substring(0, Math.Min(value.Length, maxLength));

        return value;
    }
    public static string CreateMD5(this string input)
    {
        // Use input string to calculate MD5 hash
        using (var md5 = MD5.Create())
        {
            var inputBytes = Encoding.ASCII.GetBytes(input);
            var hashBytes = md5.ComputeHash(inputBytes);

            // Convert the byte array to hexadecimal string
            var builder = new StringBuilder();
            for (var i = 0; i < hashBytes.Length; i++) builder.Append(hashBytes[i].ToString("X2").ToLower());

            return builder.ToString();
        }
    }


    /// <summary>
    ///     Extracts scheme and credential from Authorization header (if present)
    /// </summary>
    /// <param name="context"></param>
    /// <returns></returns>
    public static (string, string) GetSchemeAndCredential(HttpContext context)
    {
        var header = context.Request.Headers["Authorization"].FirstOrDefault();

        if (string.IsNullOrEmpty(header)) return (string.Empty, string.Empty);

        var parts = header.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != 2) return (string.Empty, string.Empty);

        return (parts[0], parts[1]);
    }
    public static int ParseString2Int(this string str)
    {
        int output = 0;
        Int32.TryParse(str, out output);
        return output;
    }
    public static bool ParseString2Bool(this string str)
    {
        Boolean output;
        Boolean.TryParse(str, out output);
        return output;
    }
    public static double ParseString2Double(this string value)
    {
        double outVal = 0;
        if (!String.IsNullOrEmpty(value))
        {

            double.TryParse(value, out outVal);

            if (double.IsNaN(outVal) || double.IsInfinity(outVal))
            {
                return 0;
            }
            return outVal;
        }

        return outVal;
    }

    private static readonly string[] VietnameseSigns = new string[]
        {
            "aAeEoOuUiIdDyY",
            "áàạảãâấầậẩẫăắằặẳẵ",
            "ÁÀẠẢÃÂẤẦẬẨẪĂẮẰẶẲẴ",
            "éèẹẻẽêếềệểễ",
            "ÉÈẸẺẼÊẾỀỆỂỄ",
            "óòọỏõôốồộổỗơớờợởỡ",
            "ÓÒỌỎÕÔỐỒỘỔỖƠỚỜỢỞỠ",
            "úùụủũưứừựửữ",
            "ÚÙỤỦŨƯỨỪỰỬỮ",
            "íìịỉĩ",
            "ÍÌỊỈĨ",
            "đ",
            "Đ",
            "ýỳỵỷỹ",
            "ÝỲỴỶỸ"
        };
    public static string RemoveSign4VietnameseString(string str)
    {
        for (int i = 1; i < VietnameseSigns.Length; i++)
        {
            for (int j = 0; j < VietnameseSigns[i].Length; j++)
                str = str.Replace(VietnameseSigns[i][j], VietnameseSigns[0][i - 1]);
        }
        return str;
    }

    public static string GetRealIP(this IHeaderDictionary headers, string key)
    {
        var dataHeader = headers
             .ToDictionary(item => item.Key.ToLower(), item => item.Value);

        var result = string.Empty;
        if (dataHeader.ContainsKey(key.ToLower()) == false)
        {
            return string.Empty;
        }

        result = dataHeader[key.ToLower()].NullToEmpty();

        return result;
    }

    public static string MaskString(this string source, int start, int maskLength, char maskCharacter = '*')
    {
        string mask = new string(maskCharacter, maskLength);
        string unMaskStart = source.Substring(0, start);
        string unMaskEnd = source.Substring(start + maskLength, source.Length - maskLength);
        return unMaskStart + mask + unMaskEnd;
    }

    public static T ConvertQueryStringToObject<T>(this string fromObject)
    {
        var dict = HttpUtility.ParseQueryString(fromObject);
        string json = JsonConvert.SerializeObject(dict.Cast<string>().ToDictionary(k => k, v => dict[v]));
        var result = JsonConvert.DeserializeObject<T>(json);
        return result;
    }

}

