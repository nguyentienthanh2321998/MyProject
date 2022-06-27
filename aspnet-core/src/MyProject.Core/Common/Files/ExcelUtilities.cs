
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
public static class ExcelUtilities
{
    public static Stream CreateWorkbook(
        DataSet dataset,
        bool isCreatedStream = false,
        params DataTable[] tableConfigDescription)
    {
        if (dataset.Tables.Count == 0)
            throw new ArgumentException("DataSet needs to have at least one DataTable", "dataset");

        ExcelPackage.LicenseContext = LicenseContext.Commercial;
        using var excel = new ExcelPackage();

        foreach (DataTable table in dataset.Tables)
        {
            var sheet = excel.Workbook.Worksheets.Add(table.TableName);
            sheet.Cells["A1"].LoadFromDataTable(table, true);
            sheet.Cells.SetStyleDefault(false);

            if (isCreatedStream) continue;
            var columnIndex = 1;
            foreach (DataColumn column in table.Columns)
            {
                var columnCurrent = sheet.Column(columnIndex);
                if (column.DataType == typeof(DateTime))
                    columnCurrent.Style.Numberformat.Format = "MM/dd/yyyy hh:mm:ss AM/PM";

                columnCurrent.AutoFit();

                var header = sheet.Cells[1, columnIndex];
                header.SetStyleDefault();

                if (tableConfigDescription != null
                    && tableConfigDescription.Contains(table))
                {
                    var headerDescription = sheet.Cells[2, columnIndex];
                    headerDescription.SetStyleDefault(false);
                }

                columnIndex++;
            }
        }

        var stream = new MemoryStream(excel.GetAsByteArray());

        return stream;
    }

    public static void SetStyleDefault(
        this ExcelRange range,
        bool isHeader = true,
        ExcelHorizontalAlignment horizontalAlignment = ExcelHorizontalAlignment.Left,
        ExcelVerticalAlignment verticalAlignment = ExcelVerticalAlignment.Center)
    {
        range.Style.Font.Bold = isHeader;
        range.Style.Font.Name = "Calibri";
        range.Style.Font.Size = 12;
        range.Style.HorizontalAlignment = horizontalAlignment;
        range.Style.VerticalAlignment = verticalAlignment;
        range.Style.WrapText = true;

        if (isHeader)
        {
            range.Style.Fill.SetBackground(Color.MediumSeaGreen);
            range.Style.Font.Size = 16;
            range.Style.WrapText = false;
            range.AutoFitColumns();
        }
    }

    public static int GetColumnByName(this ExcelWorksheet worksheet, string columnName)
    {
        if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));

        var columns = worksheet.Cells["1:1"]
            .Select(c => new KeyValuePair<int, string>(c.Start.Column, c.Value.NullToEmpty().ToLower()))
            .ToList();

        var column = columns.FirstOrDefault(item => item.Value == columnName.NullToEmpty().ToLower());

        var result = string.IsNullOrEmpty(column.Value) ? -1 : column.Key;

        return result;
    }

    public static string GetAddressWithHeader(
        this ExcelRangeColumn columnProperty,
        int endRow,
        bool isFixedReferenceRange = true,
        int rowAdd = 10)
    {
        endRow += rowAdd; // set thêm 10 row
        // Lấy cột A, B, C, ...
        var column = columnProperty.Range.Start.Address.Replace("1", string.Empty);
        var address = isFixedReferenceRange
            ? columnProperty.Range.FullAddress.Replace(
                columnProperty.Range.Address,
                string.Format(
                    "${0}$1:${0}${1}", // $2 => lấy row thứ 2 , bỏ qua header
                    column,
                    endRow))
            : string.Format(
                "{0}1:{0}{1}", // 2 => lấy row thứ 2 , bỏ qua header
                column,
                endRow);

        return address;
    }

    public static string GetAddressWithoutHeader(
        this ExcelRangeColumn columnProperty,
        int endRow,
        bool isFixedReferenceRange = true,
        int rowAdd = 10)
    {
        endRow += rowAdd; // set thêm 10 row
        // Lấy cột A, B, C, ...
        var column = columnProperty.Range.Start.Address.Replace("1", string.Empty);
        var address = isFixedReferenceRange
            ? columnProperty.Range.FullAddress.Replace(
                columnProperty.Range.Address,
                string.Format(
                    "${0}$2:${0}${1}", // $2 => lấy row thứ 2 , bỏ qua header
                    column,
                    endRow))
            : string.Format(
                "{0}2:{0}{1}", // 2 => lấy row thứ 2 , bỏ qua header
                column,
                endRow);

        return address;
    }

    /// <summary>
    /// Excel To Dictionary
    /// </summary>
    /// <param name="filePath">File Path</param>
    /// <param name="sheetNames">Sheet Names</param>
    /// <returns>Object Data with json</returns>
    /// <remarks>
    /// Nếu get all sheet thì không cần truyền
    /// </remarks>
    public static Dictionary<string, List<CustomDynamicObject>> ExcelToDictionary(
        string filePath,
        params string[] sheetNames)
    {
        ExcelPackage.LicenseContext = LicenseContext.Commercial;
        using var excel = new ExcelPackage(filePath);
        var data = excel.Workbook.Worksheets
            .Where(item => sheetNames.Length == 0 || sheetNames.Contains(item.Name.Trim()))
            .Select(item => item.Cells[item.Dimension.Address].ToDataTable(_ => { _.DataTableName = item.Name; }))
            .ToArray();
        var result = ExcelToDictionary(data);

        return result;
    }

    public static Dictionary<string, List<CustomDynamicObject>> ExcelToDictionary(
        FileStream filePath)
    {
        ExcelPackage.LicenseContext = LicenseContext.Commercial;
        using var excel = new ExcelPackage(filePath);
        var data = excel.Workbook.Worksheets
            .Select(item => item.Cells[item.Dimension.Address].ToDataTable(_ => { _.DataTableName = item.Name; }))
            .ToArray();
        var result = ExcelToDictionary(data);

        return result;
    }

    public static Dictionary<string, List<CustomDynamicObject>> ExcelToDictionary(DataTable[] tables)
    {
        var result = new Dictionary<string, List<CustomDynamicObject>>();

        foreach (var table in tables)
        {
            var json = JsonConvert.SerializeObject(table);
            var jsonArray = JArray.Parse(json);
            if (jsonArray == null)
            {
                result.Add(table.TableName, null);
            }
            else
            {
                var dataFilter = jsonArray.Where(item => item.HasValues);
                var dynamicObject = dataFilter
                    .Select(item => item.ToObject<CustomDynamicObject>())
                    .ToList();
                result.Add(table.TableName, dynamicObject);
            }
        }

        return result;
    }

    public static Dictionary<string, List<CustomDynamicObject>> ExcelToDictionary(
        this Stream stream)
    {
        //Instantiate the spreadsheet creation engine.
        using var excelEngine = new ExcelEngine();

        //Instantiate the Excel application object.
        var application = excelEngine.Excel;
        application.DefaultVersion = ExcelVersion.Xlsx;

        //Load the input Excel file
        var book = application.Workbooks.Open(stream);
        stream.Close();

        // Get worksheet name
        var worksheetNames = book.Worksheets
            .Select(worksheet => worksheet.Name)
            .ToArray();

        using var jsonStream = new MemoryStream();
        book.SaveAsJson(jsonStream); //Save the entire workbook as a JSON stream

        excelEngine.Dispose();

        var json = new byte[jsonStream.Length];

        //Read the Json stream and convert to a Json object
        jsonStream.Position = 0;
        jsonStream.Read(json, 0, (int)jsonStream.Length);

        // Remove object blank
        var jsonString = Encoding.UTF8.GetString(json).Replace("{    },", string.Empty);

        // Remove line blank
        var value = Regex.Replace(
            jsonString,
            @"^\s+$[\r\n]*",
            string.Empty,
            RegexOptions.Multiline);

        value = Regex.Replace(
            value,
            @"^\s+$[\r]*",
            string.Empty,
            RegexOptions.Multiline);

        value = Regex.Replace(
            value,
            @"^\s+$[\n]*",
            string.Empty,
            RegexOptions.Multiline);

        var jsonObject = JObject.Parse(value);

        var result = new Dictionary<string, List<CustomDynamicObject>>();
        foreach (var worksheetName in worksheetNames)
        {
            var jsonArray = jsonObject[worksheetName] as JArray;
            if (jsonArray == null)
            {
                result.Add(worksheetName, null);
            }
            else
            {
                var dataFilter = jsonArray.Where(item => item.HasValues);
                var dynamicObject = dataFilter
                    .Select(item => item.ToObject<CustomDynamicObject>())
                    .ToList();
                result.Add(worksheetName, dynamicObject);
            }
        }

        return result;
    }
}